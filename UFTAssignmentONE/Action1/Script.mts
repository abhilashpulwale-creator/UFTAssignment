' Initialise - Create a text file
Set fso = CreateObject("Scripting.FileSystemObject")    
Set fileOutput = fso.CreateTextFile("D:\UFT Demo\output.txt",True)

'If Browser is open,close it
If Browser("Amazon.com. Spend less.").Exist(5) Then
	Browser("Amazon.com. Spend less.").Quit
End If

'Launch Amazon
SystemUtil.Run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe","https://www.amazon.in/"
'Browser("Amazon.com. Spend less.").Sync

'Navigate to the Mobile Accessories list
Browser("Amazon.com. Spend less.").Page("Mobile Accessories: Buy").WebButton("Open All Categories Menu").WaitProperty "visible",true,30000
Browser("Amazon.com. Spend less.").Page("Mobile Accessories: Buy").WebButton("Open All Categories Menu").Click @@ script infofile_;_ZIP::ssf5.xml_;_
Browser("Amazon.com. Spend less.").Page("Mobile Accessories: Buy").Link("Mobiles, Computers").WaitProperty "visible",true,30000
Browser("Amazon.com. Spend less.").Page("Mobile Accessories: Buy").Link("Mobiles, Computers").Click @@ script infofile_;_ZIP::ssf6.xml_;_
Browser("Amazon.com. Spend less.").Page("Mobile Accessories: Buy").Link("All Mobile Accessories").Click @@ script infofile_;_ZIP::ssf7.xml_;_

'Description object to capture all links
Set objMobAccLinks = Description.Create
objMobAccLinks("xpath").value = "//li[contains(@class,'apb-browse-refinements')]//a//span"

Set ChildObj = Browser("Amazon.com. Spend less.").Page("Mobile Accessories: Buy").ChildObjects(objMobAccLinks)

'Write the mobile accessories list to file
For iCount = 0 To ChildObj.Count-1 Step 1
	fileOutput.WriteLine(ChildObj(iCount).GetROProperty("text"))
Next

'Close and clear objects
Browser("Amazon.com. Spend less.").Close
fileOutput.Close
Set fileOutput = nothing





