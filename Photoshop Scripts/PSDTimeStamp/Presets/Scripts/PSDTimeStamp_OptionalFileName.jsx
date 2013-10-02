﻿// SCRIPT NAME: PSD Time Stamp// AUTHOR: Alistair Braz// DESCRIPTION: Saves PSD with timestamp and gives option for the User to change file name.// Version: 1.0//EXCEPTION CATCH - RUNNING SCRIPT WITH NO DOCUMENTS OPENtry{    app.bringToFront();    var docRef = activeDocument;    main();}catch(e){    alert("There are no documents open.");}//MAIN FUNCTIONfunction main(){    try{        //Get time from getTime function        var time = getTime();        //Get original file name from getFileName function        var fileName = getFileName(docRef);        //Get path where original file sits and use as target location        var savePath = getPath(docRef);        //Save file        saveDoc(docRef,time,fileName,savePath);    }    catch(e){       //If fails error warning       alert("Was not able to save.");    }}//HELPER FUNCTIONSfunction getTime(){     var currentTime = new Date();     var timeOfDay;     var month = currentTime.getMonth() + 1;     var day = currentTime.getDate();     var year = currentTime.getFullYear();     var hours = currentTime.getHours();     var minutes = currentTime.getMinutes();          //fixes issues where if its 01 minute would show up as 1:1 PM for example     if (minutes < 10) {            minutes = "0" + minutes;     }          //Determine the time of day     if (hours<11) {            timeOfDay = "AM";     }     else {            timeOfDay = "PM";     }     //Make timestamp     var timeStamp = day + "-" + month + "-" + year + " " + hours + "-" + minutes + " " + timeOfDay;     return timeStamp;}function getPath(docRef) {    //Need to save before I can get the path.          docRef.save();    var savePath = docRef.path;    return savePath;}function getFileName(docRef){    var documentName = docRef.name;    //Stripping the extension out from the file name and casting into new var    var strippedExtension = documentName.split(".");    var fileName = strippedExtension[0];    return fileName;}function saveDoc(docRef,time,fileName,savePath){    var extension = ".psd";    //Building path for file    var cPath = savePath + "/" + fileName + " " + time + extension;    var saveFile = new File(cPath);    var saveOptions = new PhotoshopSaveOptions();    //Save    docRef.saveAs(saveFile, saveOptions, true,Extension.LOWERCASE);}