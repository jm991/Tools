StampedPngOnSave
John Mac (john@outlook.com)

Use:
I mainly use this guy at work when making presentations and I need to save out "states" of my Photoshop doc. By "state," I mean that I'll change a bunch of layers to get one variation of a design and save the doc - with this script enabled, a separate PNG file with the time stamped on it is created and stored alongside. Think of it as a snapshot.


Install:
Place the .jsx file in C:\Program Files\Adobe\Adobe Photoshop CC (64 Bit)\Presets\Scripts\Event Scripts Only
Open Photoshop CC:
File>Scripts>Scripts Event Manager
Tick Enable Events to Run Scripts/Actions
Select Photoshop Event "Save Document"
In the next dropdown box, select your new script and click add.

Now every time you do a save, the script will check if it is a "PSD" document, if it is it will save a timestamped PNG with the same name to the same location.