/* Given a spreadsheet such as GU_Start_End_Times_Report_AUG.xlsx exported from R25
Determine the earliest class start times (minStartTime), and latest class end times per class (maxEndTime)
Then determine minStartTime, maxEndTime per HVAC Zone
This code consumes all dates as columns in a spreadsheet, and classrooms as RTCOfferAnswerOptions
Note: As of May 2019, data will be exported from 25 Live and may look different
*/

const xlsx = require("xlsx");
const _ = require('underscore');

var wb = xlsx.readFile("GU_Start_End_Times_Room_Report_AUG.xlsx", {cellDates:true});

var ws = wb.Sheets["GU_Start_End_Times_Room_Report_"];

var data = xlsx.utils.sheet_to_json(ws)

var dLen = data.length;
var minStartTimeArray = [];
var maxEndTimeArray = [];
var minStartTime;
var maxEndTime;
var classroomMinMax = [];
var currentClassroom;


for (i=0; i < dLen; i++){ //iterate through each row of data
    var classRowCounter=0;
    for (const [key, value] of Object.entries(data[i])){ //iterate through each column for the current row[i]
        if (classRowCounter===0) { //first column is always classroom name, set to currentClassroom.
            currentClassroom = value;
            classRowCounter++;
        } else { //every other column besides the first is a date column depicting a time range for that date. split each time range into startTime - endTime
        var startTime = value.substring(0,8);
        var endTime = value.substring(17,9);
        minStartTimeArray.push(startTime);
        maxEndTimeArray.push(endTime);
        //console.log(`${startTime} - ${endTime}`);
        classRowCounter++;
        } //if classRowCounter===0 else
    } //for each Classroom column
    minStartTimeArray.sort(function (a, b) {
        return new Date('1492/01/01 ' + a) - new Date('1492/01/01 ' + b); //used to allow sorting of AM/PM values, 1492 is arbitrary date
      });
    minStartTime=minStartTimeArray[0]; //first element in minStartTimeArray is the earliest start time.
    //console.log(minStartTime);
    minStartTimeArray=[];
    maxEndTimeArray.sort(function (a, b) {
        return new Date('1492/01/01 ' + a) - new Date('1492/01/01 ' + b); //used to allow sorting of AM/PM values, 1482 is arbitrary date
      });
    maxEndTime=maxEndTimeArray[maxEndTimeArray.length-1] //last element in maxEndTimeArray is the latest end time.
    //console.log(maxStartTime);
    maxEndTimeArray=[];
    classroomMinMax.push([currentClassroom,minStartTime,maxEndTime]);
    console.log(classroomMinMax);
    
}; //for each Classroom row
    
/* var hvacZonesArray=[];
var zoneMinMaxArray=[];
var zoneMin, zoneMax;


var found = classroomMinMax.find(function(classroom) {
    return classroom[0]==="GAR 101";
});
console.log(found); */

console.log(`-=-=---=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-`)

//iterate using underscore.js instead:
/* _.each( data, function(room, key){
  _.each(room, function(roomNo,roomKey){
    console.log(roomNo)
  });
}); */

//TODO
//For each HVAC ZONE:
// Search classroomMinMax array to find classrooms in HVAC zone and determine minStartTime and maxEndTime for the group of classrooms in each zone
// For each calculated HVAC zone min/max, push that entry into zoneMinMaxArray as [classroom,zoneMin,zoneMax]