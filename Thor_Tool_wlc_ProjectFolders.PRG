******************************************************************************************
*  PROGRAM: Thor_Tool_ProjectFolders.prg
*
*  AUTHOR: Richard A. Schummer, November 2014
*
*  COPYRIGHT © 2014-2015   All Rights Reserved.
*     White Light Computing, Inc.
*     Rick Schummer
*     PO Box 391
*     Washington Twp., MI  48094
*
*     raschummer@whitelightcomputing.com
*
*  EXPLICIT LICENSE:
*     White Light Computing grants a perpetual, non-transferable, non-exclusive, 
*     royalty free, worldwide license to use and employ such materials within 
*     their business to other Visual FoxPro developers, with full derivative rights.
*
*  DISCLAIMER OF WARRANTIES. 
*     The Software is provided "AS IS" and "WITH ALL FAULTS," without warranty of any kind,
*     including without limitation the warranties of merchantability, fitness for a 
*     particular purpose and non-infringement. The Licensor makes no warranty that 
*     the Software is free of defects or is suitable for any particular purpose. In no 
*     event shall the Licensor be responsible for loss or damages arising from the installation 
*     or use of the Software, including but not limited to any indirect, punitive, special, 
*     incidental or consequential damages of any character including, without limitation, 
*     damages for loss of goodwill, work stoppage, computer failure or malfunction, or any 
*     and all other commercial damages or losses. The entire risk as to the quality and 
*     performance of the Software is borne by you. Should the Software prove defective, you 
*     and not the Licensor assume the entire cost of any service and repair.
*
*  PROGRAM DESCRIPTION:
*     This program checks to see if the data in tables can be cleanly opened.
*
*  CALLING SYNTAX:
*     DO Thor_Tool_ProjectFolders.prg
*     Thor_Tool_ProjectFolders()
*
*  INPUT PARAMETERS:
*     lxParam1 = unknown data type, typically the standard Thor object passed to register
*                the tool with Thor and appear on the Thor menu.
*
*  OUTPUT PARAMETERS:
*     None
*
*  DATABASES ACCESSED:
*     None
* 
*  GLOBAL PROCEDURES REQUIRED:
*     None
* 
*  CODING STANDARDS:
*     Version 5.2 compliant with no exceptions
*  
*  TEST INFORMATION:
*     None
*   
*  SPECIAL REQUIREMENTS/DEVICES:
*     None
*
*  FUTURE ENHANCEMENTS:
*     None
*
*  LANGUAGE/VERSION:
*     Visual FoxPro 09.00.0000.7423 or higher
* 
******************************************************************************************
*                             C H A N G E    L O G                              
*
*    Date     Developer               Version  Description
* ----------  ----------------------  -------  ---------------------------------
* 11/19/2014  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     1.1      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

IF PCOUNT() = 1 AND 'O' = VARTYPE(lxParam1) AND 'thorinfo' == LOWER(lxParam1.Class)

   WITH lxParam1
   
      * Required
      .Prompt         = 'List Project Paths'      && used in menus
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2     && a description for the tool
         This program opens up a project and lists the paths found in the project, along with the number of base source files in each folder. 
         
         You can select either full paths or relative paths via message box question. If one project is available in the current directory, it is automatically selected, otherwise you can select a project via the open dialog. 
         
         Additionally, a CSV file is created in the VFP Temp folder (SYS(2023)) with the project name and '_ProjectFolders' as the file name in case you need the information for other purposes.
      ENDTEXT  
      
      .StatusBarText = 'Display all the paths from a VFP project file.'  
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source        = "WLC"                            && where did this tool come from?  Your own initials, for instance
      .Category      = "WLC"                            && creates categorization of tools; defaults to .Source if empty
      .Sort          = 0                                && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version       = "Version 1.1, September 20, 2015"           && e.g., 'Version 7, May 18, 2011'
      .Author        = "Rick Schummer"
      .Link          = "https://github.com/rschummer/ThorTools"    && link to a page for this tool
      .VideoLink     = SPACE(0)                                    && link to a video for this tool
      
   ENDWITH 

   RETURN lxParam1
ENDIF 

IF PCOUNT() = 0
   DO ToolCode
ELSE
   DO ToolCode WITH lxParam1
ENDIF

RETURN

****************************************************************
****************************************************************
* Normal processing for this tool begins here.                  
PROCEDURE ToolCode

LPARAMETERS tcProjectName

LOCAL lnProjectFilesAvailable, ;
      lcOldSafety, ;
      lnResult, ;
      llFullPath, ;
      lcExportFile

#DEFINE ccCRLF              CHR(13) + CHR(10)

SET ASSERTS ON

* Make sure that the proper parameter(s) are passed
IF PCOUNT() = 1 AND VARTYPE(tcProjectName) = "C" AND NOT EMPTY(tcProjectName)
   ASSERT FILE(FORCEEXT(tcProjectName,"PJX")) MESSAGE "Can't find specified project"
ELSE
   lnProjectFilesAvailable = ADIR(laPjx, "*.pjx")
   
   IF lnProjectFilesAvailable # 1 
      tcProjectName = GETFILE("PJX", "Pick the project")
   ELSE
      tcProjectName = laPjx[1,1]
   ENDIF 
   
   IF EMPTY(tcProjectName)
      RETURN 
   ENDIF 
ENDIF

TRY 
   lcOldExclusive = SET("Exclusive")
   lcOldDeleted   = SET("Deleted")
   lcOldTalk      = SET("Talk")
   lcOldEscape    = SET("Escape")

   SET EXCLUSIVE OFF
   SET DELETED ON
   SET TALK OFF
   SET ESCAPE ON
   
   CloseCursors()
   
   lnResult = MESSAGEBOX("Do you want full paths included in the directory summary?", ; 
                         0+4+32 , ;
                         _screen.Caption)

   llFullPath = lnResult = 6      && Yes
   
   USE (tcProjectName) IN 0 SHARED AGAIN ALIAS curProject
   
   lcExportFile = ADDBS(SYS(2023)) + FORCEEXT(JUSTSTEM(DBF("curProject")) + "_ProjectFolders", "csv")

   IF llFullPath
      SELECT CAST(LOWER(FULLPATH(JUSTPATH(name), ADDBS(JUSTPATH(DBF("curProject"))))) AS C(250)) AS cDirectories,; 
             COUNT(*) AS nFiles ;
         FROM curProject ;
         ORDER BY 1 ;
         GROUP BY 1 ;
         INTO CURSOR curDirectories
   ELSE 
      SELECT CAST(JUSTPATH(name) AS C(250)) AS cDirectories,; 
             COUNT(*) AS nFiles ;
         FROM curProject ;
         ORDER BY 1 ;
         GROUP BY 1 ;
         INTO CURSOR curDirectories
   ENDIF
   
   SELECT curDirectories
   
   lcOldSafety = SET("Safety")
   SET SAFETY OFF 
 
   COPY TO (lcExportFile) TYPE CSV

   SET SAFETY &lcOldSafety
   
   GO TOP IN curDirectories
   BROWSE LAST NOCAPTIONS NOWAIT 
   
   * Inform developer of file to use later if they like
   WAIT WINDOW "Files exported to: " + LOWER(lcExportFile) + "..." NOWAIT
   
   * Allow the project to be opened after the process is run
   CloseProjectCursor()
   
CATCH TO loException
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(lcCode, ;
              0+48, ;
              _screen.Caption)

FINALLY
   SET EXCLUSIVE &lcOldExclusive
   SET DELETED &lcOldDeleted
   SET TALK &lcOldTalk
   SET ESCAPE &lcOldEscape 
ENDTRY

* Reset the message 
SET MESSAGE TO 
   
RETURN

ENDPROC 


********************************************************************************
*  METHOD NAME: CloseCursors
*
*  AUTHOR: Richard A. Schummer, November 2014
*
*  METHOD DESCRIPTION:
*    This method is called to close the cursors opened by the tool.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CloseCursors()

USE IN (SELECT("curDirectories"))
CloseProjectCursor()

RETURN 


********************************************************************************
*  METHOD NAME: CloseProjectCursor
*
*  AUTHOR: Richard A. Schummer, December 2014
*
*  METHOD DESCRIPTION:
*    This method is called to close the project cursor opened by the tool.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CloseProjectCursor()

USE IN (SELECT("curProject"))

RETURN 

*: EOF :*