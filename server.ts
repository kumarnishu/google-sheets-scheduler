// import task sheet
let TaskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(PropertiesService.getScriptProperties().getProperty("tasks_sheet")))

// import task triggers sheet
let TaskTriggersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(PropertiesService.getScriptProperties().getProperty("task_triggers_sheet")))

//import logger sheet
let LogsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(PropertiesService.getScriptProperties().getProperty("logs_sheet")))


//create scheduler menu
function onOpen() {
    SpreadsheetApp.getUi().createMenu("Scheduler")
        .addItem("Start Task Scheduler", 'StartTaskScheduler')
        .addSeparator()
        .addItem("Start Greeting Scheduler", 'StartGreetingScheduler')
        .addSeparator()
        .addItem("Stop Task Scheduler", 'DeleteTaskTriggers')
        .addSeparator()
        .addItem("Stop Greeting Scheduler", 'DeleteGreetingTriggers')
        .addToUi()

}


//get request or web app
function doGet(e: GoogleAppsScript.Events.DoGet) {
    const ScriptProperty = PropertiesService.getScriptProperties()
    //Displays the text on the webpage.
    let mode = e.parameter["hub.mode"];
    let challange = e.parameter["hub.challenge"];
    let token = e.parameter["hub.verify_token"];
    if (mode && token) {
        if (mode === "subscribe" && token === ScriptProperty.getProperty('myToken')) {
            return ContentService.createTextOutput(challange)
        } else {
            return ContentService.createTextOutput(JSON.stringify({ error: 'Error message' })).setMimeType(ContentService.MimeType.JSON);
        }
    }
}

//post request
function doPost(e: GoogleAppsScript.Events.DoPost) {
    const ScriptProperty = PropertiesService.getScriptProperties()
    let token = ScriptProperty.getProperty('accessToken')
    if (!token) {
        ServerLog("Please provide valid access token")
    }
    const { entry } = JSON.parse(e.postData.contents)
    ServerLog(e.postData.contents)
    try {
        if (entry.length > 0 && token) {
            if (entry[0].changes[0].value.messages) {
                let type = entry[0].changes[0].value.messages[0].type
                let from = entry[0].changes[0].value.messages[0].from
                switch (type) {
                    case "button": {
                        let btnRes = entry[0].changes[0].value.messages[0].button.text
                        let timestamp = new Date(entry[0].changes[0].value.messages[0].timestamp * 1000)
                        let wamid = String(entry[0].changes[0].value.messages[0].context.id)
                        ServerLog(btnRes + String(timestamp))
                        UpdateTaskStatus(wamid, btnRes, timestamp)
                        sendTextMessage(`Response Saved`, from, token)
                    }
                        break;
                    case "text": {
                        let response = String(entry[0].changes[0].value.messages[0].text.body)
                        let timestamp = new Date(entry[0].changes[0].value.messages[0].timestamp * 1000)
                        sendTextMessage(`Hi , We Got Your Message, We will Reply Soon`, from, token)
                    }
                        break;
                    default: sendTextMessage(`failed to parse message `, from, token)
                }
            }
            if (entry[0].changes[0].value.statuses) {
                let status = String(entry[0].changes[0].value.statuses[0].status)
                let wamid = String(entry[0].changes[0].value.statuses[0].id)
                let timestamp = new Date(entry[0].changes[0].value.statuses[0].timestamp * 1000)
                UpdateTaskWhatsappStatus(wamid, status, timestamp)
                UpdateGreetingWhatsappStatus(wamid, status, timestamp)
            }
        }
    }
    catch (error: any) {
        ServerLog(error)
        console.log(error)
    }
}

//update task whatsapp message status
function UpdateTaskWhatsappStatus(wamid: string, response: string, timestamp: Date) {
    if (TaskSheet) {
        for (let i = 3; i <= TaskSheet.getLastRow(); i++) {
            let message_id = String(TaskSheet?.getRange(i, 26).getValue())
            if (message_id === wamid) {
                TaskSheet?.getRange(i, 3).setValue(response)
                TaskSheet?.getRange(i, 4).setValue(timestamp)
            }
        }
    }
}

//update task status (pending/done)
function UpdateTaskStatus(wamid: string, response: string, timestamp: Date) {
    if (TaskSheet) {
        for (let i = 3; i <= TaskSheet.getLastRow(); i++) {
            let message_id = String(TaskSheet?.getRange(i, 26).getValue())
            if (wamid === message_id) {
                TaskSheet?.getRange(i, 5).setValue(response.toLowerCase())
                TaskSheet?.getRange(i, 6).setValue(timestamp)
                if (response.toLowerCase() === "done") {
                    TaskSheet?.getRange(i, 14).setValue("stop")
                }
            }
        }
    }
}

// trigger type to store in trigger sheet
type Trigger = {
    date: Date,
    refresh_date: Date,
    trigger_id: string,
    trigger_type: string,
    id: number,
    phone: Number,
    name: string,
    task_title: string,
    task_detail: string,
    mf: number,
    hf: number,
    df: number,
    wf: number,
    monthf: number,
    yearf: number,
    weekdays: string,
    monthdays: string
}


//start scheduler function
function StartTaskScheduler() {
    if (TaskSheet) {
        //trigger error handler
        for (let i = 3; i <= TaskSheet.getLastRow(); i++) {
            let scheduler_status = String(TaskSheet?.getRange(i, 2).getValue())
            let autoStop = String(TaskSheet?.getRange(i, 14).getValue())
            let task_status = String(TaskSheet?.getRange(i, 5).getValue())
            let phone = String(TaskSheet?.getRange(i, 10).getValue())
            if (autoStop.toLowerCase() !== "stop" && task_status.toLowerCase() !== "done" && scheduler_status.toLowerCase() !== "running" && scheduler_status.toLowerCase() !== "ready" && phone) {
                if (TaskErrorHandler(i))
                    return
            }
        }
        //setup start date trigger
        for (let i = 3; i <= TaskSheet.getLastRow(); i++) {
            let scheduler_status = String(TaskSheet?.getRange(i, 2).getValue())
            let autoStop = String(TaskSheet?.getRange(i, 14).getValue())
            let task_status = String(TaskSheet?.getRange(i, 5).getValue())
            let phone = String(TaskSheet?.getRange(i, 10).getValue())
            if (autoStop.toLowerCase() !== "stop" && task_status.toLowerCase() !== "done" && scheduler_status.toLowerCase() !== "running" && scheduler_status.toLowerCase() !== "ready" && phone) {
                SetUpTaskStartDateTrigger(i)
                TaskFirstRefreshDateUpdater(i)
            }
        }
    }
}

function TaskErrorHandler(index: number) {
    let errorStatus = false
    let date = new Date(TaskSheet?.getRange(index, 23).getValue())
    let phoneno = TaskSheet?.getRange(index, 10).getValue()
    let rowid = TaskSheet?.getRange(index, 1).getValue();
    let mf = TaskSheet?.getRange(index, 15).getValue();
    let hf = TaskSheet?.getRange(index, 16).getValue();
    let df = TaskSheet?.getRange(index, 17).getValue();
    let wf = TaskSheet?.getRange(index, 18).getValue();
    let monthf = TaskSheet?.getRange(index, 19).getValue();
    let yearf = TaskSheet?.getRange(index, 20).getValue();
    let weekdays = String(TaskSheet?.getRange(index, 21).getValue())
    let monthdays = String(TaskSheet?.getRange(index, 22).getValue())
    let mins = [0, 1, 5, 10, 15, 30]
    if (!mf) mf = 0
    if (typeof (mf) !== "number" || !mins.includes(mf)) {
        Alert(`Select valid  minutes ,choose one from 0,1,5,10,15,30: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0

    if (!isvalidDate(date)) {
        Alert(`Select valid  date: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (date < new Date()) {
        Alert(`Select valid  date ,date could not be in the past: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    let TmpArr = [mf, hf, df, wf, monthf, yearf]
    if (!phoneno) {
        Alert(`Select Phone no first : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (!rowid) {
        Alert(`Id for this row not exists : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (String(phoneno).length < 12) {
        Alert(`Select Phone no in correct format : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    let count = 0
    TmpArr.forEach((item) => {
        if (item > 0) {
            count++;
        }
    });
    let tmpWeekdays = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
    if (weekdays.length > 0) {
        let weekDays = weekdays.split(",")
        weekDays.forEach((item) => {
            if (!tmpWeekdays.includes(item.toLowerCase())) {
                Alert(`Select week days in correct format : Error comes from Row No ${index} In Data Range`)
                errorStatus = true
            }
        })
        count++
    }
    if (String(monthdays).length > 0) {
        let monthDays = monthdays.split(",")
        monthDays.forEach((item) => {
            if (Number(item) === 0 || item.length > 2 || Number(item) > 28) {
                Alert(`Select month days in correct format less than 29 and more than 0 : Error comes from Row No ${index} In Data Range`)
                errorStatus = true
            }

        })
        count++
    }
    if (count > 1) {
        Alert(`Select only one from from hour,minutes,days,weeks,year ,week days, and month days repeatation : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (errorStatus)
        return true
}

function SetUpTaskStartDateTrigger(index: number) {
    let date = new Date(TaskSheet?.getRange(index, 23).getValue())
    let refresh_date = date
    let phone = TaskSheet?.getRange(index, 10).getValue()
    let name = TaskSheet?.getRange(index, 9).getValue()
    let task_title = TaskSheet?.getRange(index, 7).getValue()
    let task_detail = TaskSheet?.getRange(index, 8).getValue()
    let id = TaskSheet?.getRange(index, 1).getValue();
    let mf = TaskSheet?.getRange(index, 15).getValue();
    let hf = TaskSheet?.getRange(index, 16).getValue();
    let df = TaskSheet?.getRange(index, 17).getValue();
    let wf = TaskSheet?.getRange(index, 18).getValue();
    let monthf = TaskSheet?.getRange(index, 19).getValue();
    let yearf = TaskSheet?.getRange(index, 20).getValue();
    let weekdays = String(TaskSheet?.getRange(index, 21).getValue())
    let monthdays = String(TaskSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let trigger = ScriptApp.newTrigger('SetUpTaskWhatsappTrigger').timeBased().at(date).create()
    SaveTaskTrigger({
        date: date,
        refresh_date: refresh_date,
        trigger_id: trigger.getUniqueId(),
        trigger_type: trigger.getHandlerFunction(),
        id: id,
        phone: phone,
        name: name,
        task_title: task_title,
        task_detail: task_detail,
        mf: mf,
        hf: hf,
        df: df,
        wf: wf,
        monthf: monthf,
        yearf: yearf,
        weekdays: weekdays,
        monthdays: monthdays
    })
    TaskSheet?.getRange(index, 2).setValue("ready").setFontWeight('bold')
}

function SetUpTaskWhatsappTrigger(e: GoogleAppsScript.Events.TimeDriven) {
    let triggers = findAllTaskTriggers().filter((trigger) => {
        if (trigger.trigger_id === e.triggerUid && trigger.trigger_type === "SetUpTaskWhatsappTrigger") {
            return trigger
        }
    })
    if (triggers.length > 0) {
        let index = findIndexOfTaskById(triggers[0].id)
        if (index)
            WhatsappTaskTrigger(index)
    }
}

function WhatsappTaskTrigger(index: number) {
    let date = new Date(TaskSheet?.getRange(index, 23).getValue())
    let refresh_date = date
    let phone = TaskSheet?.getRange(index, 10).getValue()
    let name = TaskSheet?.getRange(index, 9).getValue()
    let task_title = TaskSheet?.getRange(index, 7).getValue()
    let task_detail = TaskSheet?.getRange(index, 8).getValue()
    let id = TaskSheet?.getRange(index, 1).getValue();
    let mf = TaskSheet?.getRange(index, 15).getValue();
    let hf = TaskSheet?.getRange(index, 16).getValue();
    let df = TaskSheet?.getRange(index, 17).getValue();
    let wf = TaskSheet?.getRange(index, 18).getValue();
    let monthf = TaskSheet?.getRange(index, 19).getValue();
    let yearf = TaskSheet?.getRange(index, 20).getValue();
    let weekdays = String(TaskSheet?.getRange(index, 21).getValue())
    let monthdays = String(TaskSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let triggers: GoogleAppsScript.Script.Trigger[] = []
    if (mf > 0) {
        let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().everyMinutes(mf).create();
        triggers.push(tr)
    }
    if (hf > 0) {
        let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().everyHours(hf).create();
        triggers.push(tr)
    }
    if (df > 0) {
        let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().everyDays(df).atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
    }
    if (wf > 0) {
        let weekday = ScriptApp.WeekDay.SUNDAY
        if (date.getDay() === 1)
            weekday = ScriptApp.WeekDay.MONDAY
        if (date.getDay() === 2)
            weekday = ScriptApp.WeekDay.TUESDAY
        if (date.getDay() === 3)
            weekday = ScriptApp.WeekDay.WEDNESDAY
        if (date.getDay() === 4)
            weekday = ScriptApp.WeekDay.THURSDAY
        if (date.getDay() === 5)
            weekday = ScriptApp.WeekDay.FRIDAY
        if (date.getDay() === 6)
            weekday = ScriptApp.WeekDay.SATURDAY
        if (date.getDay() === 7)
            weekday = ScriptApp.WeekDay.SUNDAY
        let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().everyWeeks(wf).onWeekDay(weekday)
            .atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
    }
    if (monthf > 0) {
        let totaldaystoadd = GetMonthDays(date.getFullYear(), date.getMonth()) - date.getDate()
        for (let i = 0; i < Number(monthf); i++) {
            date = new Date(date.setMonth(date.getMonth() + 1))
            totaldaystoadd += GetMonthDays(date.getFullYear(), date.getMonth())
        }
        let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().everyDays(totaldaystoadd).atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
    }
    if (yearf > 0) {
        let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().everyDays(GetYearDays(date.getFullYear()) * yearf).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
        triggers.push(tr)
    }
    if (weekdays.length > 0) {
        weekdays.split(",").forEach((wd) => {
            if (wd.toLowerCase() === "sun") {
                let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "mon") {
                let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "tue") {
                let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "wed") {
                let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "thu") {
                let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "fri") {
                let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "sat") {
                let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
        })
    }
    if (monthdays.length > 0) {
        monthdays.split(",").forEach((md) => {
            let tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().onMonthDay(Number(md)).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
            triggers.push(tr)
        })
    }

    let tr: GoogleAppsScript.Script.Trigger | undefined = undefined
    tr = ScriptApp.newTrigger('SendTaskMessage').timeBased().at(date).create();
    if (tr)
        triggers.push(tr)
    triggers.forEach((trigger) => {
        SaveTaskTrigger({
            date: date,
            refresh_date: refresh_date,
            trigger_id: trigger.getUniqueId(),
            trigger_type: trigger.getHandlerFunction(),
            id: id,//track trigger with excel rows
            phone: phone,
            name: name,
            task_title: task_title,
            task_detail: task_detail,
            mf: mf,
            hf: hf,
            df: df,
            wf: wf,
            monthf: monthf,
            yearf: yearf,
            weekdays: weekdays,
            monthdays: monthdays
        })
    })

    if (index)
        TaskSheet?.getRange(index, 2).setValue("running").setFontWeight('bold')
    
}

function TaskFirstRefreshDateUpdater(index: number) {
    let date = new Date(TaskSheet?.getRange(index, 23).getValue())
    let refresh_date = date
    let mf = TaskSheet?.getRange(index, 15).getValue();
    let hf = TaskSheet?.getRange(index, 16).getValue();
    let df = TaskSheet?.getRange(index, 17).getValue();
    let wf = TaskSheet?.getRange(index, 18).getValue();
    let monthf = TaskSheet?.getRange(index, 19).getValue();
    let yearf = TaskSheet?.getRange(index, 20).getValue();
    let weekdays = String(TaskSheet?.getRange(index, 21).getValue())
    let monthdays = String(TaskSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0

    if (mf === 1) {
        let miliseconds = 30000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (mf > 1) {
        let miliseconds = 3 * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (hf > 0) {
        let miliseconds = 10 * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (df > 0) {
        let miliseconds = 30 * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (wf > 0) {
        let miliseconds = (30) * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (monthf > 0) {
        let miliseconds = (30) * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (yearf > 0) {
        let miliseconds = (30) * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (weekdays.length > 0) {
        let miliseconds = 30 * 60000
        refresh_date = new Date(refresh_date.getTime() - miliseconds)

    }
    if (monthdays.length > 0) {
        let miliseconds = 30 * 60000
        refresh_date = new Date(refresh_date.getTime() - miliseconds)
    }
    TaskSheet?.getRange(index, 24).setValue(new Date(refresh_date))
    SetUpTaskRefreshDateTrigger(index, refresh_date)
}

function SetUpTaskRefreshDateTrigger(index: number, refresh_date: Date) {
    let date = new Date(TaskSheet?.getRange(index, 23).getValue())
    let phone = TaskSheet?.getRange(index, 10).getValue()
    let name = TaskSheet?.getRange(index, 9).getValue()
    let task_title = TaskSheet?.getRange(index, 7).getValue()
    let task_detail = TaskSheet?.getRange(index, 8).getValue()
    let id = TaskSheet?.getRange(index, 1).getValue();
    let mf = TaskSheet?.getRange(index, 15).getValue();
    let hf = TaskSheet?.getRange(index, 16).getValue();
    let df = TaskSheet?.getRange(index, 17).getValue();
    let wf = TaskSheet?.getRange(index, 18).getValue();
    let monthf = TaskSheet?.getRange(index, 19).getValue();
    let yearf = TaskSheet?.getRange(index, 20).getValue();
    let weekdays = String(TaskSheet?.getRange(index, 21).getValue())
    let monthdays = String(TaskSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let triggers: GoogleAppsScript.Script.Trigger[] = []

    if (mf > 0) {
        let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().everyMinutes(mf).create();
        triggers.push(tr)
    }
    if (hf > 0) {
        let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().everyHours(hf).create();
        triggers.push(tr)
    }
    if (df > 0) {
        let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().everyDays(df).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create()
        triggers.push(tr)
    }
    if (wf > 0) {
        let weekday = ScriptApp.WeekDay.SUNDAY
        if (refresh_date.getDay() === 1)
            weekday = ScriptApp.WeekDay.MONDAY
        if (refresh_date.getDay() === 2)
            weekday = ScriptApp.WeekDay.TUESDAY
        if (refresh_date.getDay() === 3)
            weekday = ScriptApp.WeekDay.WEDNESDAY
        if (refresh_date.getDay() === 4)
            weekday = ScriptApp.WeekDay.THURSDAY
        if (refresh_date.getDay() === 5)
            weekday = ScriptApp.WeekDay.FRIDAY
        if (refresh_date.getDay() === 6)
            weekday = ScriptApp.WeekDay.SATURDAY
        if (refresh_date.getDay() === 7)
            weekday = ScriptApp.WeekDay.SUNDAY
        let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().everyWeeks(wf).onWeekDay(weekday)
            .atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create()
        triggers.push(tr)
    }
    if (monthf > 0) {
        let totaldaystoadd = GetMonthDays(date.getFullYear(), date.getMonth()) - date.getDate()
        for (let i = 0; i < Number(monthf); i++) {
            date = new Date(date.setMonth(date.getMonth() + 1))
            totaldaystoadd += GetMonthDays(date.getFullYear(), date.getMonth())
        }
        let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().everyDays(totaldaystoadd).atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
    }

    if (yearf > 0) {
        let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().everyDays(GetYearDays(date.getFullYear()) * yearf).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
        triggers.push(tr)
    }
    if (weekdays.length > 0) {
        weekdays.split(",").forEach((wd) => {
            if (wd.toLowerCase() === "sun") {
                let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "mon") {
                let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "tue") {
                let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "wed") {
                let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "thu") {
                let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "fri") {
                let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "sat") {
                let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
        })
    }
    if (monthdays.length > 0) {
        let tr = ScriptApp.newTrigger('TaskRefreshDateTrigger').timeBased().onMonthDay(Number(monthdays.split(",")[0])).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
        triggers.push(tr)
    }
    triggers.forEach((trigger) => {
        SaveTaskTrigger({
            date: date,
            refresh_date: refresh_date,
            trigger_id: trigger.getUniqueId(),
            trigger_type: trigger.getHandlerFunction(),
            id: id,//track trigger with excel rows
            phone: phone,
            name: name,
            task_title: task_title,
            task_detail: task_detail,
            mf: mf,
            hf: hf,
            df: df,
            wf: wf,
            monthf: monthf,
            yearf: yearf,
            weekdays: weekdays,
            monthdays: monthdays
        })
    })
    if (index)
        TaskSheet?.getRange(index, 2).setValue("running").setFontWeight('bold')
}
function TaskRefreshDateTrigger(e: GoogleAppsScript.Events.TimeDriven) {
    let triggers = findAllTaskTriggers().filter((trigger) => {
        if (trigger.trigger_id === String(e.triggerUid) && trigger.trigger_type === "TaskRefreshDateTrigger") {
            return trigger
        }
    })
    if (triggers.length > 0) {
        let index = findIndexOfTaskById(triggers[0].id)
        if (index) {
            TaskRefreshDateUpdater(index)
        }
    }
}
function TaskRefreshDateUpdater(index: number) {
    let refresh_date = new Date(TaskSheet?.getRange(index, 24).getValue())
    let date = new Date(TaskSheet?.getRange(index, 23).getValue())
    let mf = TaskSheet?.getRange(index, 15).getValue();
    let hf = TaskSheet?.getRange(index, 16).getValue();
    let df = TaskSheet?.getRange(index, 17).getValue();
    let wf = TaskSheet?.getRange(index, 18).getValue();
    let monthf = TaskSheet?.getRange(index, 19).getValue();
    let yearf = TaskSheet?.getRange(index, 20).getValue();
    let weekdays = String(TaskSheet?.getRange(index, 21).getValue())
    let monthdays = String(TaskSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    if (mf > 0) {
        let miliseconds = mf * 60000
        refresh_date = new Date(refresh_date.getTime() + miliseconds)
    }
    if (hf > 0) {
        let miliseconds = hf * 60 * 60000
        refresh_date = new Date(refresh_date.getTime() + miliseconds)
    }
    if (df > 0) {
        refresh_date = new Date(refresh_date.setDate(refresh_date.getDate() + Number(df)))
    }
    if (wf > 0) {
        refresh_date = new Date(refresh_date.setDate(refresh_date.getDate() + Number(wf * 7)))
    }
    if (monthf > 0) {
        refresh_date = new Date(refresh_date.setMonth(refresh_date.getMonth() + monthf));
    }
    if (yearf > 0) {
        refresh_date = new Date(refresh_date.setFullYear(refresh_date.getFullYear() + yearf));
    }

    if (weekdays.length > 0) {
        let current_day = String(refresh_date).split(" ")[0]
        let daysgap = 0
        if (current_day.toLowerCase() === "sun") {
            daysgap = daysgap + 7
        }
        if (current_day.toLowerCase() === "mon") {
            daysgap = daysgap + 6
        }
        if (current_day.toLowerCase() === "tue") {
            daysgap = daysgap + 5
        }
        if (current_day.toLowerCase() === "wed") {
            daysgap = daysgap + 4
        }
        if (current_day.toLowerCase() === "thu") {
            daysgap = daysgap + 3
        }
        if (current_day.toLowerCase() === "fri") {
            daysgap = daysgap + 2
        }
        if (current_day.toLowerCase() === "sat") {
            daysgap = daysgap + 1
        }
        let wd = weekdays.split(",")[0].toLowerCase()
        if (wd === "mon")
            daysgap = daysgap + 1
        if (wd === "tue")
            daysgap = daysgap + 2
        if (wd === "wed")
            daysgap = daysgap + 3
        if (wd === "thu")
            daysgap = daysgap + 4
        if (wd === "fri")
            daysgap = daysgap + 5
        if (wd === "sat")
            daysgap = daysgap + 6
        refresh_date = new Date(refresh_date.setDate(refresh_date.getDate() + daysgap))
    }
    if (monthdays.length > 0) {
        if (Number(monthdays.split(",")[0]) < date.getDate())
            refresh_date = new Date(date.setMonth(date.getMonth() + 1))
        let miliseconds = 30 * 60000
        refresh_date = new Date(refresh_date.setDate(Number(monthdays.split(",")[0])))
        refresh_date = new Date(refresh_date.getTime() - miliseconds)
    }
    TaskSheet?.getRange(index, 24).setValue(refresh_date)
    TaskSheet?.getRange(index, 14).setValue("")
    TaskSheet?.getRange(index, 5).setValue("pending")
}

function TaskLastDateUpdater(index: number) {
    let date = new Date(TaskSheet?.getRange(index, 23).getValue())
    let mf = TaskSheet?.getRange(index, 15).getValue();
    let hf = TaskSheet?.getRange(index, 16).getValue();
    let df = TaskSheet?.getRange(index, 17).getValue();
    let wf = TaskSheet?.getRange(index, 18).getValue();
    let monthf = TaskSheet?.getRange(index, 19).getValue();
    let yearf = TaskSheet?.getRange(index, 20).getValue();
    let weekdays = String(TaskSheet?.getRange(index, 21).getValue())
    let monthdays = String(TaskSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    if (mf > 0) {
        let miliseconds = mf * 60000
        date = new Date(date.getTime() + miliseconds)
    }
    if (hf > 0) {
        let miliseconds = hf * 60 * 60000
        date = new Date(date.getTime() + miliseconds)
    }
    if (df > 0) {
        date = new Date(date.setDate(date.getDate() + Number(df)))
    }
    if (wf > 0) {
        date = new Date(date.setDate(date.getDate() + Number(wf * 7)))
    }
    if (monthf > 0) {
        date = new Date(date.setMonth(date.getMonth() + monthf));
    }
    if (yearf > 0) {
        date = new Date(date.setFullYear(date.getFullYear() + yearf));
    }
    if (weekdays.length > 0) {
        let current_day = String(date).split(" ")[0]
        let daysgap = 0
        if (current_day.toLowerCase() === "sun") {
            daysgap = daysgap + 7
        }
        if (current_day.toLowerCase() === "mon") {
            daysgap = daysgap + 6
        }
        if (current_day.toLowerCase() === "tue") {
            daysgap = daysgap + 5
        }
        if (current_day.toLowerCase() === "wed") {
            daysgap = daysgap + 4
        }
        if (current_day.toLowerCase() === "thu") {
            daysgap = daysgap + 3
        }
        if (current_day.toLowerCase() === "fri") {
            daysgap = daysgap + 2
        }
        if (current_day.toLowerCase() === "sat") {
            daysgap = daysgap + 1
        }
        let wd = weekdays.split(",")[0].toLowerCase()
        if (wd === "mon")
            daysgap = daysgap + 1
        if (wd === "tue")
            daysgap = daysgap + 2
        if (wd === "wed")
            daysgap = daysgap + 3
        if (wd === "thu")
            daysgap = daysgap + 4
        if (wd === "fri")
            daysgap = daysgap + 5
        if (wd === "sat")
            daysgap = daysgap + 6
        date = new Date(date.setDate(date.getDate() + daysgap))
    }
    if (monthdays.length > 0) {
        if (Number(monthdays.split(",")[0]) < date.getDate())
            date = new Date(date.setMonth(date.getMonth() + 1))
        date = new Date(date.setDate(Number(monthdays.split(",")[0])))
    }

    TaskSheet?.getRange(index, 11).setValue(date)
}

function SaveTaskTrigger(trigger: Trigger) {
    let row = TaskTriggersSheet?.getDataRange().getLastRow()
    if (row)
        row = row + 1
    if (row) {
        TaskTriggersSheet?.getRange(row, 1).setValue(trigger.date)
        TaskTriggersSheet?.getRange(row, 2).setValue(trigger.refresh_date)
        TaskTriggersSheet?.getRange(row, 3).setValue(trigger.trigger_id)
        TaskTriggersSheet?.getRange(row, 4).setValue(trigger.trigger_type)
        TaskTriggersSheet?.getRange(row, 5).setValue(trigger.id)
        TaskTriggersSheet?.getRange(row, 6).setValue(trigger.phone)
        TaskTriggersSheet?.getRange(row, 7).setValue(trigger.name)
        TaskTriggersSheet?.getRange(row, 8).setValue(trigger.task_title)
        TaskTriggersSheet?.getRange(row, 9).setValue(trigger.task_detail)
        TaskTriggersSheet?.getRange(row, 10).setValue(trigger.mf)
        TaskTriggersSheet?.getRange(row, 11).setValue(trigger.hf)
        TaskTriggersSheet?.getRange(row, 12).setValue(trigger.df)
        TaskTriggersSheet?.getRange(row, 13).setValue(trigger.wf)
        TaskTriggersSheet?.getRange(row, 14).setValue(trigger.monthf)
        TaskTriggersSheet?.getRange(row, 15).setValue(trigger.yearf)
        TaskTriggersSheet?.getRange(row, 16).setValue(trigger.weekdays)
        TaskTriggersSheet?.getRange(row, 17).setValue(trigger.monthdays)
    }
}
function TaskAutoStop(id: number) {
    let stop = false
    if (TaskSheet) {
        for (let i = 3; i <= TaskSheet.getLastRow(); i++) {
            let scheduler_status = String(TaskSheet?.getRange(i, 2).getValue())
            let row_id = Number(TaskSheet?.getRange(i, 1).getValue())
            let autostop = String(TaskSheet?.getRange(i, 14).getValue())
            if (scheduler_status.toLowerCase() === "running") {
                if (row_id === id) {
                    if (autostop.toLowerCase() === "stop") {
                        stop = true
                    }
                }
            }
        }
    }
    return stop
}
function DeleteTaskTriggers() {
    let delete_row_ids: number[] = []
    let delete_row_ids_in_task_trigger_sheet: number[] = []
    if (TaskSheet) {
        for (let i = 3; i <= TaskSheet.getLastRow(); i++) {
            let rowid = Number(TaskSheet?.getRange(i, 1).getValue())
            let is_delete = Boolean(TaskSheet?.getRange(i, 25).getValue())
            if (is_delete) {
                delete_row_ids.push(rowid)
                TaskSheet?.getRange(i, 2).clearContent()
                TaskSheet?.getRange(i, 3).clearContent()
                TaskSheet?.getRange(i, 4).clearContent()
                TaskSheet?.getRange(i, 5).clearContent()
                TaskSheet?.getRange(i, 6).clearContent()
                TaskSheet?.getRange(i, 14).clearContent()
                TaskSheet?.getRange(i, 24).clearContent()
                TaskSheet?.getRange(i, 26).clearContent()
            }
        }
    }
    if (TaskTriggersSheet) {
        for (let i = 2; i <= TaskTriggersSheet.getLastRow(); i++) {
            let rowid = Number(TaskTriggersSheet?.getRange(i, 5).getValue())
            let tr_id = String(TaskTriggersSheet?.getRange(i, 3).getValue())
            if (delete_row_ids.includes(rowid)) {
                delete_row_ids_in_task_trigger_sheet.push(i)
                TaskTriggersSheet.deleteRow(i)
                DeleteTrigger(tr_id)
                i--
            }
        }
    }
}
function DeleteTrigger(trigger_id: String) {
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (trigger.getUniqueId() === trigger_id) {
            ScriptApp.deleteTrigger(trigger)
        }
    })
}
function deleteTaskTriggerFromTriggerSheet(index: number) {
    TaskTriggersSheet?.deleteRow(index)
}
function findIndexOfTaskById(id: number) {
    if (TaskSheet) {
        for (let i = 3; i <= TaskSheet.getLastRow(); i++) {
            let rowid = Number(TaskSheet?.getRange(i, 1).getValue())
            if (rowid === id) {
                return i
            }
        }
    }
}
function SetTaskMessageId(id: number, message_id: String) {
    let index = findIndexOfTaskById(id)
    if (index) {
        if (TaskSheet) {
            TaskSheet.getRange(index, 26).setValue(message_id)
        }
    }
}
function GetMonthDays(year: number, month: number) {
    let febDays = 28
    if (year % 4 === 0) {
        febDays = 29
    }
    let day31 = [1, 3, 5, 7, 8, 10, 12]
    let day30 = [4, 6, 9, 11]
    if (day31.includes(month))
        return 31
    if (day30.includes(month))
        return 30
    return febDays
}

function GetYearDays(year: number) {
    let yeardays = 365
    if (year % 4 === 0)
        yeardays = 366
    return yeardays
}
function findAllTaskTriggers() {
    let triggers: Trigger[] = []
    if (TaskTriggersSheet) {
        for (let i = 2; i <= TaskTriggersSheet.getLastRow(); i++) {
            let date = TaskTriggersSheet?.getRange(i, 1).getValue()
            let refresh_date = TaskTriggersSheet?.getRange(i, 2).getValue()
            let trigger_id = TaskTriggersSheet?.getRange(i, 3).getValue()
            let trigger_type = TaskTriggersSheet?.getRange(i, 4).getValue()
            let id = Number(TaskTriggersSheet?.getRange(i, 5).getValue())
            let phone = TaskTriggersSheet?.getRange(i, 6).getValue()
            let name = TaskTriggersSheet?.getRange(i, 7).getValue()
            let task_title = TaskTriggersSheet?.getRange(i, 8).getValue()
            let task_detail = TaskTriggersSheet?.getRange(i, 9).getValue()
            let mf = TaskTriggersSheet?.getRange(i, 10).getValue()
            let hf = TaskTriggersSheet?.getRange(i, 11).getValue()
            let df = TaskTriggersSheet?.getRange(i, 12).getValue()
            let wf = TaskTriggersSheet?.getRange(i, 13).getValue()
            let monthf = TaskTriggersSheet?.getRange(i, 14).getValue()
            let yearf = TaskTriggersSheet?.getRange(i, 15).getValue()
            let weekdays = TaskTriggersSheet?.getRange(i, 16).getValue()
            let monthdays = TaskTriggersSheet?.getRange(i, 17).getValue()
            triggers.push({
                date: date,
                refresh_date: refresh_date,
                trigger_id: trigger_id,
                trigger_type: trigger_type,
                id: id,
                phone: phone,
                name: name,
                task_title: task_title,
                task_detail: task_detail,
                mf: mf,
                hf: hf,
                df: df,
                wf: wf,
                monthf: monthf,
                yearf: yearf,
                weekdays: weekdays,
                monthdays: monthdays
            })

        }
    }
    return triggers
}

function sendTextMessage(message: string, from: string, token: string) {
    let phone_id = PropertiesService.getScriptProperties().getProperty('phone_id')
    if (!phone_id) {
        ServerLog("provide a valid phone id")
    }
    let url = `https://graph.facebook.com/v16.0/${phone_id}/messages`;
    let data = {
        "messaging_product": "whatsapp",
        "to": from,
        "type": "text",
        "text": {
            "body": message
        }
    }
    let options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        "method": "post",
        "headers": {
            "Authorization": `Bearer ${token}`
        },
        "contentType": "application/json",
        "payload": JSON.stringify(data)
    };
    UrlFetchApp.fetch(url, options)
}
function SendTaskMessage(e: GoogleAppsScript.Events.TimeDriven) {
    let phone_id = PropertiesService.getScriptProperties().getProperty('phone_id')
    let triggers = findAllTaskTriggers().filter((trigger) => {
        if (trigger.trigger_id === e.triggerUid && trigger.trigger_type === "SendTaskMessage") {
            return trigger
        }
    })

    if (triggers.length > 0) {

        if (!TaskAutoStop(triggers[0].id)) {
            try {
                let token = PropertiesService.getScriptProperties().getProperty('accessToken')
                let url = `https://graph.facebook.com/v16.0/${phone_id}/messages`;
                let data = {
                    "messaging_product": "whatsapp",
                    "recipient_type": "individual",
                    "to": triggers[0].phone,
                    "type": "template",
                    "template": {
                        "name": "scheduler_with_response_",
                        "language": {
                            "code": "en_US"
                        },
                        "components": [
                            {
                                "type": "header",
                                "parameters": [
                                    {
                                        "type": "text",
                                        "text": triggers[0].task_title
                                    }
                                ]
                            },
                            {
                                "type": "body",
                                "parameters": [
                                    {
                                        "type": "text",
                                        "text": triggers[0].task_detail
                                    }
                                ]
                            }
                        ]
                    }
                }
                let options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
                    "method": "post",
                    "headers": {
                        "Authorization": `Bearer ${token}`
                    },
                    "contentType": "application/json",
                    "payload": JSON.stringify(data)
                };

                let response = UrlFetchApp.fetch(url, options)
                const { messages } = JSON.parse(response.getContentText())
                if (messages.length > 0) {
                    SetTaskMessageId(triggers[0].id, messages[0].id)
                }
            }
            catch (err) {
                console.log(err)
            }
        }
        let index = findIndexOfTaskById(triggers[0].id)
        if (index) {
            TaskLastDateUpdater(index)
        }
    }
}


//greeting code here

// ############################
// ###########################################
// import greeting sheet
let GreetingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(PropertiesService.getScriptProperties().getProperty("greetings_sheet")))

// import greeting triggers sheet
let GreetingTriggersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(PropertiesService.getScriptProperties().getProperty("greeting_triggers_sheet")))


//update greeting whatsapp message status
function UpdateGreetingWhatsappStatus(wamid: string, response: string, timestamp: Date) {
    if (GreetingSheet) {
        for (let i = 3; i <= GreetingSheet.getLastRow(); i++) {
            let message_id = String(GreetingSheet?.getRange(i, 26).getValue())
            if (message_id === wamid) {
                GreetingSheet?.getRange(i, 3).setValue(response)
                GreetingSheet?.getRange(i, 4).setValue(timestamp)
            }
        }
    }
}


// trigger type to store in trigger sheet
type GreetingTrigger = {
    date: Date,
    refresh_date: Date,
    trigger_id: string,
    trigger_type: string,
    id: number,
    phone: Number,
    name: string,
    greeting_image: string,
    greeting_detail: string,
    mf: number,
    hf: number,
    df: number,
    wf: number,
    monthf: number,
    yearf: number,
    weekdays: string,
    monthdays: string
}


//start scheduler function
function StartGreetingScheduler() {
    if (GreetingSheet) {
        //trigger error handler
        for (let i = 3; i <= GreetingSheet.getLastRow(); i++) {
            let scheduler_status = String(GreetingSheet?.getRange(i, 2).getValue())
            let autoStop = String(GreetingSheet?.getRange(i, 14).getValue())
            let greeting_status = String(GreetingSheet?.getRange(i, 5).getValue())
            let phone = String(GreetingSheet?.getRange(i, 10).getValue())
            if (autoStop.toLowerCase() !== "stop" && greeting_status.toLowerCase() !== "done" && scheduler_status.toLowerCase() !== "running" && scheduler_status.toLowerCase() !== "ready" && phone) {
                if (GreetingErrorHandler(i))
                    return
            }
        }
        //setup start date trigger
        for (let i = 3; i <= GreetingSheet.getLastRow(); i++) {
            let scheduler_status = String(GreetingSheet?.getRange(i, 2).getValue())
            let autoStop = String(GreetingSheet?.getRange(i, 14).getValue())
            let greeting_status = String(GreetingSheet?.getRange(i, 5).getValue())
            let phone = String(GreetingSheet?.getRange(i, 10).getValue())
            if (autoStop.toLowerCase() !== "stop" && greeting_status.toLowerCase() !== "done" && scheduler_status.toLowerCase() !== "running" && scheduler_status.toLowerCase() !== "ready" && phone) {
                SetUpGreetingStartDateTrigger(i)
                GreetingFirstRefreshDateUpdater(i)
            }
        }
    }
}

function GreetingErrorHandler(index: number) {
    let errorStatus = false
    let date = new Date(GreetingSheet?.getRange(index, 23).getValue())
    let phoneno = GreetingSheet?.getRange(index, 10).getValue()
    let rowid = GreetingSheet?.getRange(index, 1).getValue();
    let mf = GreetingSheet?.getRange(index, 15).getValue();
    let hf = GreetingSheet?.getRange(index, 16).getValue();
    let df = GreetingSheet?.getRange(index, 17).getValue();
    let wf = GreetingSheet?.getRange(index, 18).getValue();
    let monthf = GreetingSheet?.getRange(index, 19).getValue();
    let yearf = GreetingSheet?.getRange(index, 20).getValue();
    let weekdays = String(GreetingSheet?.getRange(index, 21).getValue())
    let monthdays = String(GreetingSheet?.getRange(index, 22).getValue())
    let mins = [0, 1, 5, 10, 15, 30]
    if (!mf) mf = 0
    if (typeof (mf) !== "number" || !mins.includes(mf)) {
        Alert(`Select valid  minutes ,choose one from 0,1,5,10,15,30: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0

    if (!isvalidDate(date)) {
        Alert(`Select valid  date: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (date < new Date()) {
        Alert(`Select valid  date ,date could not be in the past: Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    let TmpArr = [mf, hf, df, wf, monthf, yearf]
    if (!phoneno) {
        Alert(`Select Phone no first : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (!rowid) {
        Alert(`Id for this row not exists : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (String(phoneno).length < 12) {
        Alert(`Select Phone no in correct format : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    let count = 0
    TmpArr.forEach((item) => {
        if (item > 0) {
            count++;
        }
    });
    let tmpWeekdays = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
    if (weekdays.length > 0) {
        let weekDays = weekdays.split(",")
        weekDays.forEach((item) => {
            if (!tmpWeekdays.includes(item.toLowerCase())) {
                Alert(`Select week days in correct format : Error comes from Row No ${index} In Data Range`)
                errorStatus = true
            }
        })
        count++
    }
    if (String(monthdays).length > 0) {
        let monthDays = monthdays.split(",")
        monthDays.forEach((item) => {
            if (Number(item) === 0 || item.length > 2 || Number(item) > 28) {
                Alert(`Select month days in correct format less than 29 and more than 0 : Error comes from Row No ${index} In Data Range`)
                errorStatus = true
            }

        })
        count++
    }
    if (count > 1) {
        Alert(`Select only one from from hour,minutes,days,weeks,year ,week days, and month days repeatation : Error comes from Row No ${index} In Data Range`)
        errorStatus = true
    }
    if (errorStatus)
        return true
}

function SetUpGreetingStartDateTrigger(index: number) {
    let date = new Date(GreetingSheet?.getRange(index, 23).getValue())
    let refresh_date = date
    let phone = GreetingSheet?.getRange(index, 10).getValue()
    let name = GreetingSheet?.getRange(index, 9).getValue()
    let greeting_image = GreetingSheet?.getRange(index, 7).getValue()
    let greeting_detail = GreetingSheet?.getRange(index, 8).getValue()
    let id = GreetingSheet?.getRange(index, 1).getValue();
    let mf = GreetingSheet?.getRange(index, 15).getValue();
    let hf = GreetingSheet?.getRange(index, 16).getValue();
    let df = GreetingSheet?.getRange(index, 17).getValue();
    let wf = GreetingSheet?.getRange(index, 18).getValue();
    let monthf = GreetingSheet?.getRange(index, 19).getValue();
    let yearf = GreetingSheet?.getRange(index, 20).getValue();
    let weekdays = String(GreetingSheet?.getRange(index, 21).getValue())
    let monthdays = String(GreetingSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let trigger = ScriptApp.newTrigger('SetUpGreetingWhatsappTrigger').timeBased().at(date).create()
    SaveGreetingTrigger({
        date: date,
        refresh_date: refresh_date,
        trigger_id: trigger.getUniqueId(),
        trigger_type: trigger.getHandlerFunction(),
        id: id,
        phone: phone,
        name: name,
        greeting_image: greeting_image,
        greeting_detail: greeting_detail,
        mf: mf,
        hf: hf,
        df: df,
        wf: wf,
        monthf: monthf,
        yearf: yearf,
        weekdays: weekdays,
        monthdays: monthdays
    })
    GreetingSheet?.getRange(index, 2).setValue("ready").setFontWeight('bold')
}

function SetUpGreetingWhatsappTrigger(e: GoogleAppsScript.Events.TimeDriven) {
    let triggers = findAllGreetingTriggers().filter((trigger) => {
        if (trigger.trigger_id === e.triggerUid && trigger.trigger_type === "SetUpGreetingWhatsappTrigger") {
            return trigger
        }
    })
    if (triggers.length > 0) {
        let index = findIndexOfGreetingById(triggers[0].id)
        if (index)
            WhatsappGreetingTrigger(index)
    }
}

function WhatsappGreetingTrigger(index: number) {
    let date = new Date(GreetingSheet?.getRange(index, 23).getValue())
    let refresh_date = date
    let phone = GreetingSheet?.getRange(index, 10).getValue()
    let name = GreetingSheet?.getRange(index, 9).getValue()
    let greeting_image = GreetingSheet?.getRange(index, 7).getValue()
    let greeting_detail = GreetingSheet?.getRange(index, 8).getValue()
    let id = GreetingSheet?.getRange(index, 1).getValue();
    let mf = GreetingSheet?.getRange(index, 15).getValue();
    let hf = GreetingSheet?.getRange(index, 16).getValue();
    let df = GreetingSheet?.getRange(index, 17).getValue();
    let wf = GreetingSheet?.getRange(index, 18).getValue();
    let monthf = GreetingSheet?.getRange(index, 19).getValue();
    let yearf = GreetingSheet?.getRange(index, 20).getValue();
    let weekdays = String(GreetingSheet?.getRange(index, 21).getValue())
    let monthdays = String(GreetingSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let triggers: GoogleAppsScript.Script.Trigger[] = []
    if (mf > 0) {
        let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().everyMinutes(mf).create();
        triggers.push(tr)
       
    }
    if (hf > 0) {
        let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().everyHours(hf).create();
        triggers.push(tr)
       
    }
    if (df > 0) {
        let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().everyDays(df).atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
       
    }
    if (wf > 0) {
        let weekday = ScriptApp.WeekDay.SUNDAY
        if (date.getDay() === 1)
            weekday = ScriptApp.WeekDay.MONDAY
        if (date.getDay() === 2)
            weekday = ScriptApp.WeekDay.TUESDAY
        if (date.getDay() === 3)
            weekday = ScriptApp.WeekDay.WEDNESDAY
        if (date.getDay() === 4)
            weekday = ScriptApp.WeekDay.THURSDAY
        if (date.getDay() === 5)
            weekday = ScriptApp.WeekDay.FRIDAY
        if (date.getDay() === 6)
            weekday = ScriptApp.WeekDay.SATURDAY
        if (date.getDay() === 7)
            weekday = ScriptApp.WeekDay.SUNDAY
        let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().everyWeeks(wf).onWeekDay(weekday)
            .atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
       
    }
    if (monthf > 0) {
        let totaldaystoadd = GetMonthDays(date.getFullYear(), date.getMonth()) - date.getDate()
        for (let i = 0; i < Number(monthf); i++) {
            date = new Date(date.setMonth(date.getMonth() + 1))
            totaldaystoadd += GetMonthDays(date.getFullYear(), date.getMonth())
        }
        let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().everyDays(totaldaystoadd).atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
       
    }
    if (yearf > 0) {
        let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().everyDays(GetYearDays(date.getFullYear()) * yearf).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
        triggers.push(tr)
       
    }
    if (weekdays.length > 0) {
        weekdays.split(",").forEach((wd) => {
            if (wd.toLowerCase() === "sun") {
                let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "mon") {
                let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "tue") {
                let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "wed") {
                let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "thu") {
                let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "fri") {
                let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "sat") {
                let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
                triggers.push(tr)

            }
        })
       
    }
    if (monthdays.length > 0) {
        monthdays.split(",").forEach((md) => {
            let tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().onMonthDay(Number(md)).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
            triggers.push(tr)
        })
       
    }

    let tr: GoogleAppsScript.Script.Trigger | undefined = undefined
    tr = ScriptApp.newTrigger('SendGreetingMessage').timeBased().at(date).create();
    if (tr)
        triggers.push(tr)
    triggers.forEach((trigger) => {
        SaveGreetingTrigger({
            date: date,
            refresh_date: refresh_date,
            trigger_id: trigger.getUniqueId(),
            trigger_type: trigger.getHandlerFunction(),
            id: id,//track trigger with excel rows
            phone: phone,
            name: name,
            greeting_image: greeting_image,
            greeting_detail: greeting_detail,
            mf: mf,
            hf: hf,
            df: df,
            wf: wf,
            monthf: monthf,
            yearf: yearf,
            weekdays: weekdays,
            monthdays: monthdays
        })
    })

    if (index)
        GreetingSheet?.getRange(index, 2).setValue("running").setFontWeight('bold')
    
}

function GreetingFirstRefreshDateUpdater(index: number) {
    let date = new Date(GreetingSheet?.getRange(index, 23).getValue())
    let refresh_date = date
    let mf = GreetingSheet?.getRange(index, 15).getValue();
    let hf = GreetingSheet?.getRange(index, 16).getValue();
    let df = GreetingSheet?.getRange(index, 17).getValue();
    let wf = GreetingSheet?.getRange(index, 18).getValue();
    let monthf = GreetingSheet?.getRange(index, 19).getValue();
    let yearf = GreetingSheet?.getRange(index, 20).getValue();
    let weekdays = String(GreetingSheet?.getRange(index, 21).getValue())
    let monthdays = String(GreetingSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0

    if (mf === 1) {
        let miliseconds = 30000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (mf > 1) {
        let miliseconds = 3 * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (hf > 0) {
        let miliseconds = 10 * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (df > 0) {
        let miliseconds = 30 * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (wf > 0) {
        let miliseconds = (30) * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (monthf > 0) {
        let miliseconds = (30) * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (yearf > 0) {
        let miliseconds = (30) * 60000
        refresh_date = new Date(date.getTime() - miliseconds)
    }
    if (weekdays.length > 0) {
        let miliseconds = 30 * 60000
        refresh_date = new Date(refresh_date.getTime() - miliseconds)

    }
    if (monthdays.length > 0) {
        let miliseconds = 30 * 60000
        refresh_date = new Date(refresh_date.getTime() - miliseconds)
    }
    GreetingSheet?.getRange(index, 24).setValue(new Date(refresh_date))
    SetUpGreetingRefreshDateTrigger(index, refresh_date)
}

function SetUpGreetingRefreshDateTrigger(index: number, refresh_date: Date) {
    let date = new Date(GreetingSheet?.getRange(index, 23).getValue())
    let phone = GreetingSheet?.getRange(index, 10).getValue()
    let name = GreetingSheet?.getRange(index, 9).getValue()
    let greeting_image = GreetingSheet?.getRange(index, 7).getValue()
    let greeting_detail = GreetingSheet?.getRange(index, 8).getValue()
    let id = GreetingSheet?.getRange(index, 1).getValue();
    let mf = GreetingSheet?.getRange(index, 15).getValue();
    let hf = GreetingSheet?.getRange(index, 16).getValue();
    let df = GreetingSheet?.getRange(index, 17).getValue();
    let wf = GreetingSheet?.getRange(index, 18).getValue();
    let monthf = GreetingSheet?.getRange(index, 19).getValue();
    let yearf = GreetingSheet?.getRange(index, 20).getValue();
    let weekdays = String(GreetingSheet?.getRange(index, 21).getValue())
    let monthdays = String(GreetingSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    let triggers: GoogleAppsScript.Script.Trigger[] = []

    if (mf > 0) {
        let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().everyMinutes(mf).create();
        triggers.push(tr)
    }
    if (hf > 0) {
        let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().everyHours(hf).create();
        triggers.push(tr)
    }
    if (df > 0) {
        let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().everyDays(df).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create()
        triggers.push(tr)
    }
    if (wf > 0) {
        let weekday = ScriptApp.WeekDay.SUNDAY
        if (refresh_date.getDay() === 1)
            weekday = ScriptApp.WeekDay.MONDAY
        if (refresh_date.getDay() === 2)
            weekday = ScriptApp.WeekDay.TUESDAY
        if (refresh_date.getDay() === 3)
            weekday = ScriptApp.WeekDay.WEDNESDAY
        if (refresh_date.getDay() === 4)
            weekday = ScriptApp.WeekDay.THURSDAY
        if (refresh_date.getDay() === 5)
            weekday = ScriptApp.WeekDay.FRIDAY
        if (refresh_date.getDay() === 6)
            weekday = ScriptApp.WeekDay.SATURDAY
        if (refresh_date.getDay() === 7)
            weekday = ScriptApp.WeekDay.SUNDAY
        let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().everyWeeks(wf).onWeekDay(weekday)
            .atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create()
        triggers.push(tr)
    }
    if (monthf > 0) {
        let totaldaystoadd = GetMonthDays(date.getFullYear(), date.getMonth()) - date.getDate()
        for (let i = 0; i < Number(monthf); i++) {
            date = new Date(date.setMonth(date.getMonth() + 1))
            totaldaystoadd += GetMonthDays(date.getFullYear(), date.getMonth())
        }
        let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().everyDays(totaldaystoadd).atHour(date.getHours()).nearMinute(date.getMinutes()).create()
        triggers.push(tr)
    }

    if (yearf > 0) {
        let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().everyDays(GetYearDays(date.getFullYear()) * yearf).atHour(date.getHours()).nearMinute(date.getMinutes()).create();
        triggers.push(tr)
    }
    if (weekdays.length > 0) {
        weekdays.split(",").forEach((wd) => {
            if (wd.toLowerCase() === "sun") {
                let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)
            }
            if (wd.toLowerCase() === "mon") {
                let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "tue") {
                let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "wed") {
                let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "thu") {
                let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "fri") {
                let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
            if (wd.toLowerCase() === "sat") {
                let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
                triggers.push(tr)

            }
        })
    }
    if (monthdays.length > 0) {
        let tr = ScriptApp.newTrigger('GreetingRefreshDateTrigger').timeBased().onMonthDay(Number(monthdays.split(",")[0])).atHour(refresh_date.getHours()).nearMinute(refresh_date.getMinutes()).create();
        triggers.push(tr)
    }
    triggers.forEach((trigger) => {
        SaveGreetingTrigger({
            date: date,
            refresh_date: refresh_date,
            trigger_id: trigger.getUniqueId(),
            trigger_type: trigger.getHandlerFunction(),
            id: id,//track trigger with excel rows
            phone: phone,
            name: name,
            greeting_image: greeting_image,
            greeting_detail: greeting_detail,
            mf: mf,
            hf: hf,
            df: df,
            wf: wf,
            monthf: monthf,
            yearf: yearf,
            weekdays: weekdays,
            monthdays: monthdays
        })
    })
    if (index)
        GreetingSheet?.getRange(index, 2).setValue("running").setFontWeight('bold')
}
function GreetingRefreshDateTrigger(e: GoogleAppsScript.Events.TimeDriven) {
    let triggers = findAllGreetingTriggers().filter((trigger) => {
        if (trigger.trigger_id === String(e.triggerUid) && trigger.trigger_type === "GreetingRefreshDateTrigger") {
            return trigger
        }
    })
    if (triggers.length > 0) {
        let index = findIndexOfGreetingById(triggers[0].id)
        if (index) {
            GreetingRefreshDateUpdater(index)
        }
    }
}
function GreetingRefreshDateUpdater(index: number) {
    let refresh_date = new Date(GreetingSheet?.getRange(index, 24).getValue())
    let date = new Date(GreetingSheet?.getRange(index, 23).getValue())
    let mf = GreetingSheet?.getRange(index, 15).getValue();
    let hf = GreetingSheet?.getRange(index, 16).getValue();
    let df = GreetingSheet?.getRange(index, 17).getValue();
    let wf = GreetingSheet?.getRange(index, 18).getValue();
    let monthf = GreetingSheet?.getRange(index, 19).getValue();
    let yearf = GreetingSheet?.getRange(index, 20).getValue();
    let weekdays = String(GreetingSheet?.getRange(index, 21).getValue())
    let monthdays = String(GreetingSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    if (mf > 0) {
        let miliseconds = mf * 60000
        refresh_date = new Date(refresh_date.getTime() + miliseconds)
    }
    if (hf > 0) {
        let miliseconds = hf * 60 * 60000
        refresh_date = new Date(refresh_date.getTime() + miliseconds)
    }
    if (df > 0) {
        refresh_date = new Date(refresh_date.setDate(refresh_date.getDate() + Number(df)))
    }
    if (wf > 0) {
        refresh_date = new Date(refresh_date.setDate(refresh_date.getDate() + Number(wf * 7)))
    }
    if (monthf > 0) {
        refresh_date = new Date(refresh_date.setMonth(refresh_date.getMonth() + monthf));
    }
    if (yearf > 0) {
        refresh_date = new Date(refresh_date.setFullYear(refresh_date.getFullYear() + yearf));
    }

    if (weekdays.length > 0) {
        let current_day = String(refresh_date).split(" ")[0]
        let daysgap = 0
        if (current_day.toLowerCase() === "sun") {
            daysgap = daysgap + 7
        }
        if (current_day.toLowerCase() === "mon") {
            daysgap = daysgap + 6
        }
        if (current_day.toLowerCase() === "tue") {
            daysgap = daysgap + 5
        }
        if (current_day.toLowerCase() === "wed") {
            daysgap = daysgap + 4
        }
        if (current_day.toLowerCase() === "thu") {
            daysgap = daysgap + 3
        }
        if (current_day.toLowerCase() === "fri") {
            daysgap = daysgap + 2
        }
        if (current_day.toLowerCase() === "sat") {
            daysgap = daysgap + 1
        }
        let wd = weekdays.split(",")[0].toLowerCase()
        if (wd === "mon")
            daysgap = daysgap + 1
        if (wd === "tue")
            daysgap = daysgap + 2
        if (wd === "wed")
            daysgap = daysgap + 3
        if (wd === "thu")
            daysgap = daysgap + 4
        if (wd === "fri")
            daysgap = daysgap + 5
        if (wd === "sat")
            daysgap = daysgap + 6
        refresh_date = new Date(refresh_date.setDate(refresh_date.getDate() + daysgap))
    }
    if (monthdays.length > 0) {
        if (Number(monthdays.split(",")[0]) < date.getDate())
            refresh_date = new Date(date.setMonth(date.getMonth() + 1))
        let miliseconds = 30 * 60000
        refresh_date = new Date(refresh_date.setDate(Number(monthdays.split(",")[0])))
        refresh_date = new Date(refresh_date.getTime() - miliseconds)
    }
    GreetingSheet?.getRange(index, 24).setValue(refresh_date)
    GreetingSheet?.getRange(index, 14).setValue("")
    GreetingSheet?.getRange(index, 5).setValue("pending")
}

function GreetingLastDateUpdater(index: number) {
    let date = new Date(GreetingSheet?.getRange(index, 23).getValue())
    let mf = GreetingSheet?.getRange(index, 15).getValue();
    let hf = GreetingSheet?.getRange(index, 16).getValue();
    let df = GreetingSheet?.getRange(index, 17).getValue();
    let wf = GreetingSheet?.getRange(index, 18).getValue();
    let monthf = GreetingSheet?.getRange(index, 19).getValue();
    let yearf = GreetingSheet?.getRange(index, 20).getValue();
    let weekdays = String(GreetingSheet?.getRange(index, 21).getValue())
    let monthdays = String(GreetingSheet?.getRange(index, 22).getValue())
    if (!mf || typeof (mf) !== "number") mf = 0
    if (!hf || typeof (hf) !== "number") hf = 0
    if (!df || typeof (df) !== "number") df = 0
    if (!wf || typeof (wf) !== "number") wf = 0
    if (!monthf || typeof (monthf) !== "number") monthf = 0
    if (!yearf || typeof (yearf) !== "number") yearf = 0
    if (mf > 0) {
        let miliseconds = mf * 60000
        date = new Date(date.getTime() + miliseconds)
    }
    if (hf > 0) {
        let miliseconds = hf * 60 * 60000
        date = new Date(date.getTime() + miliseconds)
    }
    if (df > 0) {
        date = new Date(date.setDate(date.getDate() + Number(df)))
    }
    if (wf > 0) {
        date = new Date(date.setDate(date.getDate() + Number(wf * 7)))
    }
    if (monthf > 0) {
        date = new Date(date.setMonth(date.getMonth() + monthf));
    }
    if (yearf > 0) {
        date = new Date(date.setFullYear(date.getFullYear() + yearf));
    }
    if (weekdays.length > 0) {
        let current_day = String(date).split(" ")[0]
        let daysgap = 0
        if (current_day.toLowerCase() === "sun") {
            daysgap = daysgap + 7
        }
        if (current_day.toLowerCase() === "mon") {
            daysgap = daysgap + 6
        }
        if (current_day.toLowerCase() === "tue") {
            daysgap = daysgap + 5
        }
        if (current_day.toLowerCase() === "wed") {
            daysgap = daysgap + 4
        }
        if (current_day.toLowerCase() === "thu") {
            daysgap = daysgap + 3
        }
        if (current_day.toLowerCase() === "fri") {
            daysgap = daysgap + 2
        }
        if (current_day.toLowerCase() === "sat") {
            daysgap = daysgap + 1
        }
        let wd = weekdays.split(",")[0].toLowerCase()
        if (wd === "mon")
            daysgap = daysgap + 1
        if (wd === "tue")
            daysgap = daysgap + 2
        if (wd === "wed")
            daysgap = daysgap + 3
        if (wd === "thu")
            daysgap = daysgap + 4
        if (wd === "fri")
            daysgap = daysgap + 5
        if (wd === "sat")
            daysgap = daysgap + 6
        date = new Date(date.setDate(date.getDate() + daysgap))
    }
    if (monthdays.length > 0) {
        if (Number(monthdays.split(",")[0]) < date.getDate())
            date = new Date(date.setMonth(date.getMonth() + 1))
        date = new Date(date.setDate(Number(monthdays.split(",")[0])))
    }

    GreetingSheet?.getRange(index, 11).setValue(date)
}

function SaveGreetingTrigger(trigger: GreetingTrigger) {
    let row = GreetingTriggersSheet?.getDataRange().getLastRow()
    if (row)
        row = row + 1
    if (row) {
        GreetingTriggersSheet?.getRange(row, 1).setValue(trigger.date)
        GreetingTriggersSheet?.getRange(row, 2).setValue(trigger.refresh_date)
        GreetingTriggersSheet?.getRange(row, 3).setValue(trigger.trigger_id)
        GreetingTriggersSheet?.getRange(row, 4).setValue(trigger.trigger_type)
        GreetingTriggersSheet?.getRange(row, 5).setValue(trigger.id)
        GreetingTriggersSheet?.getRange(row, 6).setValue(trigger.phone)
        GreetingTriggersSheet?.getRange(row, 7).setValue(trigger.name)
        GreetingTriggersSheet?.getRange(row, 8).setValue(trigger.greeting_image)
        GreetingTriggersSheet?.getRange(row, 9).setValue(trigger.greeting_detail)
        GreetingTriggersSheet?.getRange(row, 10).setValue(trigger.mf)
        GreetingTriggersSheet?.getRange(row, 11).setValue(trigger.hf)
        GreetingTriggersSheet?.getRange(row, 12).setValue(trigger.df)
        GreetingTriggersSheet?.getRange(row, 13).setValue(trigger.wf)
        GreetingTriggersSheet?.getRange(row, 14).setValue(trigger.monthf)
        GreetingTriggersSheet?.getRange(row, 15).setValue(trigger.yearf)
        GreetingTriggersSheet?.getRange(row, 16).setValue(trigger.weekdays)
        GreetingTriggersSheet?.getRange(row, 17).setValue(trigger.monthdays)
    }
}
function GreetingAutoStop(id: number) {
    let stop = false
    if (GreetingSheet) {
        for (let i = 3; i <= GreetingSheet.getLastRow(); i++) {
            let scheduler_status = String(GreetingSheet?.getRange(i, 2).getValue())
            let row_id = Number(GreetingSheet?.getRange(i, 1).getValue())
            let autostop = String(GreetingSheet?.getRange(i, 14).getValue())
            if (scheduler_status.toLowerCase() === "running") {
                if (row_id === id) {
                    if (autostop.toLowerCase() === "stop") {
                        stop = true
                    }
                }
            }
        }
    }
    return stop
}
function DeleteGreetingTriggers() {
    let delete_row_ids: number[] = []
    let delete_row_ids_in_greeting_trigger_sheet: number[] = []
    if (GreetingSheet) {
        for (let i = 3; i <= GreetingSheet.getLastRow(); i++) {
            let rowid = Number(GreetingSheet?.getRange(i, 1).getValue())
            let is_delete = Boolean(GreetingSheet?.getRange(i, 25).getValue())
            if (is_delete) {
                delete_row_ids.push(rowid)
                GreetingSheet?.getRange(i, 2).clearContent()
                GreetingSheet?.getRange(i, 3).clearContent()
                GreetingSheet?.getRange(i, 4).clearContent()
                GreetingSheet?.getRange(i, 5).clearContent()
                GreetingSheet?.getRange(i, 6).clearContent()
                GreetingSheet?.getRange(i, 14).clearContent()
                GreetingSheet?.getRange(i, 24).clearContent()
                GreetingSheet?.getRange(i, 26).clearContent()
            }
        }
    }
    if (GreetingTriggersSheet) {
        for (let i = 2; i <= GreetingTriggersSheet.getLastRow(); i++) {
            let rowid = Number(GreetingTriggersSheet?.getRange(i, 5).getValue())
            let tr_id = String(GreetingTriggersSheet?.getRange(i, 3).getValue())
            if (delete_row_ids.includes(rowid)) {
                delete_row_ids_in_greeting_trigger_sheet.push(i)
                GreetingTriggersSheet.deleteRow(i)
                DeleteTrigger(tr_id)
                i--
            }
        }
    }
}
function deleteGreetingTriggerFromTriggerSheet(index: number) {
    GreetingTriggersSheet?.deleteRow(index)
}
function findIndexOfGreetingById(id: number) {
    if (GreetingSheet) {
        for (let i = 3; i <= GreetingSheet.getLastRow(); i++) {
            let rowid = Number(GreetingSheet?.getRange(i, 1).getValue())
            if (rowid === id) {
                return i
            }
        }
    }
}
function SetGreetingMessageId(id: number, message_id: String) {
    let index = findIndexOfGreetingById(id)
    if (index) {
        if (GreetingSheet) {
            GreetingSheet.getRange(index, 26).setValue(message_id)
        }
    }
}


function findAllGreetingTriggers() {
    let triggers: GreetingTrigger[] = []
    if (GreetingTriggersSheet) {
        for (let i = 2; i <= GreetingTriggersSheet.getLastRow(); i++) {
            let date = GreetingTriggersSheet?.getRange(i, 1).getValue()
            let refresh_date = GreetingTriggersSheet?.getRange(i, 2).getValue()
            let trigger_id = GreetingTriggersSheet?.getRange(i, 3).getValue()
            let trigger_type = GreetingTriggersSheet?.getRange(i, 4).getValue()
            let id = Number(GreetingTriggersSheet?.getRange(i, 5).getValue())
            let phone = GreetingTriggersSheet?.getRange(i, 6).getValue()
            let name = GreetingTriggersSheet?.getRange(i, 7).getValue()
            let greeting_image = GreetingTriggersSheet?.getRange(i, 8).getValue()
            let greeting_detail = GreetingTriggersSheet?.getRange(i, 9).getValue()
            let mf = GreetingTriggersSheet?.getRange(i, 10).getValue()
            let hf = GreetingTriggersSheet?.getRange(i, 11).getValue()
            let df = GreetingTriggersSheet?.getRange(i, 12).getValue()
            let wf = GreetingTriggersSheet?.getRange(i, 13).getValue()
            let monthf = GreetingTriggersSheet?.getRange(i, 14).getValue()
            let yearf = GreetingTriggersSheet?.getRange(i, 15).getValue()
            let weekdays = GreetingTriggersSheet?.getRange(i, 16).getValue()
            let monthdays = GreetingTriggersSheet?.getRange(i, 17).getValue()
            triggers.push({
                date: date,
                refresh_date: refresh_date,
                trigger_id: trigger_id,
                trigger_type: trigger_type,
                id: id,
                phone: phone,
                name: name,
                greeting_image: greeting_image,
                greeting_detail: greeting_detail,
                mf: mf,
                hf: hf,
                df: df,
                wf: wf,
                monthf: monthf,
                yearf: yearf,
                weekdays: weekdays,
                monthdays: monthdays
            })

        }
    }
    return triggers
}

function SendGreetingMessage(e: GoogleAppsScript.Events.TimeDriven) {
    let phone_id = PropertiesService.getScriptProperties().getProperty('phone_id')
    let triggers = findAllGreetingTriggers().filter((trigger) => {
        if (trigger.trigger_id === e.triggerUid && trigger.trigger_type === "SendGreetingMessage") {
            return trigger
        }
    })

    if (triggers.length > 0) {

        if (!GreetingAutoStop(triggers[0].id)) {
            try {
                let token = PropertiesService.getScriptProperties().getProperty('accessToken')
                let url = `https://graph.facebook.com/v16.0/${phone_id}/messages`;
                let data = {
                    "messaging_product": "whatsapp",
                    "recipient_type": "individual",
                    "to": triggers[0].phone,
                    "type": "template",
                    "template": {
                        "name": "work_and_greetings",
                        "language": {
                            "code": "en_US"
                        },
                        "components": [
                            {
                                "type": "header",
                                "parameters": [
                                    {
                                        "type": "image",
                                        "image": { "link": triggers[0].greeting_image }
                                    }
                                ]
                            },
                            {
                                "type": "body",
                                "parameters": [
                                    {
                                        "type": "text",
                                        "text": triggers[0].greeting_detail
                                    }
                                ]
                            }
                        ]
                    }
                }
                let options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
                    "method": "post",
                    "headers": {
                        "Authorization": `Bearer ${token}`
                    },
                    "contentType": "application/json",
                    "payload": JSON.stringify(data)
                };

                let response = UrlFetchApp.fetch(url, options)
                const { messages } = JSON.parse(response.getContentText())
                if (messages.length > 0) {
                    SetGreetingMessageId(triggers[0].id, messages[0].id)
                }
            }
            catch (err) {
                console.log(err)
            }
        }
        let index = findIndexOfGreetingById(triggers[0].id)
        if (index) {
            GreetingLastDateUpdater(index)
        }
    }
}






// utils
function isvalidDate(d: any) {
    if (Object.prototype.toString.call(d) === "[object Date]") {
        // it is a date
        if (isNaN(d)) { // d.getTime() or d.valueOf() will also work
            return false
        } else {
            return true
        }
    } else {
        return false
    }
}

function Alert(message: string) {
    SpreadsheetApp.getUi().alert(message);
    return;
}
function ServerLog(msg: string) {
    if (LogsSheet) {
        let row = LogsSheet.getLastRow() + 1
        LogsSheet?.getRange(row, 1).setValue(new Date().toLocaleString())
        LogsSheet?.getRange(row, 2).setValue(msg)
    }
}



