/*
Author: Jeremy Wray, FM&T Vision 
Version 1.0
*/

//create objects
var objArgs = WScript.Arguments;
var objFSO = WScript.CreateObject("Scripting.FileSystemObject");
var objS = WScript.CreateObject("WScript.Shell");
//app file name 
var strAppFileName = "RenameWithDate.js";
//app name and version
// 1.0 first version rename a file(s) with a timestamp
var strAppName = "RenameWithDate 1.0";
// usage instructions
var deployMsg = ".\nDeploy by right-clicking folders/files you want to backup and selecting '"
//boolean for overwriting existing files/folders
var blnOverwrite = false;

//check someone's right-clicked a file, if not see if they want to install the app
if (objArgs.length == 0){

        var blnInstall = objS.Popup("Install " + strAppName + "?",-1,strAppName,36);      
        
        //if yes
        if (blnInstall == 6){
      
            //get current location of app file
            var source_obj = objS.CurrentDirectory + "\\" + strAppFileName;       
            //get location of user's Send To menu 
            var target_obj = objS.SpecialFolders("SendTo") + "\\";
            //boolean for overwriting existing app
             var blnOverwrite = false;
           
            //function to try copying the file 
            tryCopyApp = function() {                
                try {
                     objFSO.CopyFile (source_obj, target_obj, blnOverwrite);
                }
                 //if already exists get user confirmation else quit
                catch (e) {    
                    if (e.message == "File already exists"){           
                    blnOverwrite = objS.Popup("App already exists and will be overwritten\nProceed?",-1,strAppName,49);
                        if (blnOverwrite == 1){
                            tryCopyApp(); return;         
                            }else{
                            WScript.Quit(); 
                        }                        
                    }else{
                     //show any other errors and show retry button else quit
                    if (objS.Popup(e.message,-1,strAppName,21) == 4){   
                            tryCopyApp(); return;       
                        }else{
                            WScript.Quit();
                        }
                    }       
                }
                //if we get here install was successful
                objS.Popup(strAppName + " installed to " + target_obj + deployMsg + strAppFileName + "' from the Send To menu.",-1,strAppName,64);
            }
                            
            tryCopyApp();
            
         }         

     WScript.Quit();
};

function fixDate(str){ // replace spaces with -
	return str.replace(/\W/g,'-');
}	

//get datetime stamp
var date = fixDate(new Date().toLocaleString());

folder = objFSO.GetParentFolderName(objArgs(0)); // location of folders/file(s) same for all
 
// MoveFile is used to rename a file by moving it to the same location with a new name:
 
 //loop through passed args (files)
for(var i=0; i<objArgs.length; i++){   
   var item = objArgs(i);   
     if (objFSO.FolderExists(item)){//if argument is a folder
	 	var foldertomove = objFSO.GetFolder(item).Path; //get full file path/name 
		var parentfolder = objFSO.GetParentFolderName(item)+"\\"; //get parent folders
		var foldertomovename = foldertomove.replace(parentfolder,''); //get this folder
		var folderrenamed = folder + "\\" + foldertomovename + "[" + date + "]"; // new name/path (path is same) of folder with datestamp   
		objFSO.MoveFolder(foldertomove,folderrenamed);	//do the move	 	
	 }else{ //if argument is a file
	 	var filetomove = objFSO.GetFile(item).Path; //get full file path/name 
	    var filetomovename = objFSO.GetFileName(item); //get just filename
	    var filetomovextn = (filetomovename.match(/\..*/)) ? filetomovename.match(/\..*/) : '' ; //get file extn
	    var filetomovename = (filetomovextn) ? filetomovename.replace(filetomovextn,'') : filetomovename; // filename without extension 
	    var filerenamed = folder + "\\" + filetomovename + "[" + date + "]" + filetomovextn; // new name/path (path is same) of file with datestamp   
	    objFSO.MoveFile(filetomove,filerenamed);	//do the move
	 }
} 
WScript.Quit(); 