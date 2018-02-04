OfficeBrute
====================

Why did you make this?
---------------------
I was angry at a malicious document because it was password protected. So I made a thing to brute force the password by using a wordlist.  You could try it on this file for example
https://www.hybrid-analysis.com/sample/7b6c00c1a9ec9aaf5538b52ad3736515bd81e97b7bf0f5a3e8dc2e2836f59dc7


 [[https://github.com/olanderofc/OfficeBrute/blob/master/officebrute.png|alt=octocat]]

## Warning

I have not looked into how the communication using C# with word documents are impacting macros or exploits. It may be possible that you execute code which means that you should never run OfficeBrute on a live system. Only use it with malicious samples where you are in control. Microsoft is not entirely clear on how everything works.

You have been warned :)

I take no responsibility for how this program is used. My hope is that others see it as a good tool to use and perhaps add more stuff to it.

## Technical details
Using the C# Office libraries mentioned here https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.document.aspx
I will explain more in detailed what is going on by using some code copied from the project. 

````
// Setup Microsoft Word connection
Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();

// Define the null value
object nullobj = System.Reflection.Missing.Value;

// Create the document object
Microsoft.Office.Interop.Word.Document aDoc = null;
wordApp.Visible = false;

 // Open the document
 aDoc = wordApp.Documents.Open(ref FName, ref nullobj, ref nullobj, ref nullobj, ref password, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);

 //Check if the document is protected
 if (aDoc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
 {
     MessageBox.Show("Document is not protected. Will not proceed");
 }

 try
 {
  // Assign the current password to the passwd object
  object passwd = line;
  //Try to unprotect the document. If the password is wrong, an exception will be thrown. This is how we "brute force" it.
  aDoc.Unprotect(ref passwd);
  break;
 }
 catch (Exception)
 {
  // We catch the exception but do not care about it
  counter++;
  continue;
 }

````

