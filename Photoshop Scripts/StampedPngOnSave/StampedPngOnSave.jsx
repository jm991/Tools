main();

function main()
{
	var Name = app.activeDocument.name.replace(/\.[^\.]+$/, '');
	var Ext = decodeURI(app.activeDocument.name).replace(/^.*\./,'');
	if (Ext.toLowerCase() != 'psd') return;
	var Path = app.activeDocument.path;
	var saveFile = File(Path + "/" + Name + " " + getTime() + ".png");
	if (saveFile.exists) saveFile.remove();
	SavePNG(saveFile);
}

function SavePNG(saveFile)
{
	pngSaveOptions = new PNGSaveOptions();
	activeDocument.saveAs(saveFile, pngSaveOptions, true, Extension.LOWERCASE);
} 

function getTime()
{
	var currentTime = new Date();
	var timeOfDay;
	var month = currentTime.getMonth() + 1;
	var day = currentTime.getDate();
	var year = currentTime.getFullYear();
	var hours = currentTime.getHours();
	var minutes = currentTime.getMinutes();
	var seconds = currentTime.getSeconds();

	// fixes issues where if its 01 minute would show up as 1:1 PM for example
	if (minutes < 10) 
	{
		minutes = "0" + minutes;
	}
	
	if (seconds < 10) 
	{
		seconds = "0" + seconds;
	}

	// Determine the time of day
	if (hours < 11) 
	{
		timeOfDay = "AM";
	}
	else 
	{
		timeOfDay = "PM";
	}

	// Make timestamp
	var timeStamp = day + "-" + month + "-" + year + " " + hours + "." + minutes + "." + seconds + "  " + timeOfDay;
	return timeStamp;
}