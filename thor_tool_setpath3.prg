********************************************************************************
*  PROGRAM: SetPath.prg
*
*  AUTHOR: Richard A. Schummer, April 2001
*
*  COPYRIGHT © 2001-2015   All Rights Reserved.
*     Richard A. Schummer
*     White Light Computing, Inc.
*     PO Box 391
*     Washington Twp., MI  48094
*
*     raschummer@whitelightcomputing.com
*     rick@rickschummer.com
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
*     This program opens up a project and sets the appropriate pathing
*     to the files in the project. Allows forms to be run standalone
*     and compiles to take place without worry if the paths are set correctly,
*     or the files used by the project are in the project (if you take that approach),
*     or have generic tools you use and have those is various directories.
*
*     Instructions:
*        - Copy PRG to subfolder of Thor: Thor\Thor\Tools\My Tools
*        - Restart Thor.
*        - Adds menu, Thor Tools > White Light Computing > Set Path to Project Paths
*        - Set hotkey so I can run it quickly.
*
*     If a project is open and active, the program prompts developer to see if they 
*     want to get pathing from the project. 
*       - If Yes, project is closed and paths are set.
*       - If No, program checks current folder and sees how many projects are available.
*             * If one, this project is opened without a prompt
*             * If more than one or none, developer is prompted to pick the project.
*
*    If the project is open in another VFP session, developer is notified the project 
*    cannot be opened and the pathing is not set. 
*
*    What paths are set?
*       - The VFP folder is added to the path automatically, as are a couple of 
*         WLC development folders if they exist
*       - The program checks for common developer subfolders to add to the path
*       - Any other folder for a file in the project is added to the path
*       - Paths are shortened to be relative to the project folder to avoid SET PATH limitation
*    
*    Programs are used to build SET PROCEDURE
*    Class Libraries are used to build SET CLASSLIB
*
*    A log file is created when you use the program. It logs when it starts and how long 
*    it takes to set things up. If there are problems with the files it will log those 
*    issues too. Optionally, you can log the settings of SET PATH, SET CLASSLIB and 
*    SET PROCEDURE to the log file. See clLOGPATHS constant in the beginning of the 
*    main procedure. You control the log file name and folder location using the two 
*    constants: ccSETPATHERRLOGPATH and ccSETPATHERRLOGFILE
*
*    The same paths can be shown on the VFP Desktop when the program finishes figuring 
*    everything out. See clSHOWPATHS constant in the beginning of the ToolCode procedure.
*
*    Extensibility:
*    Since this program builds the paths generically, it is nice to have some hooks into 
*    customizing the program for your needs. To accomplish this, there are some procedures
*    you can add code to customize the paths set up. Check out the following procs:
*       SetPathDeveloperPreferenceDirectories(): add your folder preferences
*       SetPathCommonProjectDirectories(): add standard project directories to the path. 
*                                          Preset to common ones. This boosts performance 
*                                          for bigger projects.
*       AddCommonNonProjectClassLibs(): SET CLASSLIB to any non-project classlibs here
*       AddCommonSetProcedures(): SET PROCEDURE TO any non-project classlibs here
*
*  CALLING SYNTAX:
*     DO SetPath WITH txParam1
*     SetPath(txParam1)
*
*  INPUT PARAMETERS:
*     txParam1 = unknown type, optional, standard Thor object passed during startup
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
*     Version 5.0 compliant with no exceptions
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
*     Visual FoxPro 09.00.0000 or higher
* 
********************************************************************************
*                             C H A N G E    L O G                              
*
*    Date     Developer               Version  Description
* ----------------------------------------------------------------------------------------
* 01/04/2001  Richard A. Schummer     1.0      Adopted program created at 
*                                              Kirtland Associates
* ----------------------------------------------------------------------------------------
* 04/04/2009  Richard A. Schummer     1.1      Cleanup and added local error handling for
*                                              duplicate names in SET CLASSLIB
* ----------------------------------------------------------------------------------------
* 10/04/2014  Richard A. Schummer     2.0      Adopted for Thor integration
* ----------------------------------------------------------------------------------------
* 11/13/2014  Richard A. Schummer     2.1      Minor clean up of the code
* ----------------------------------------------------------------------------------------
* 06/17/2015  Richard A. Schummer     2.2      Clean up of the code, added ability to
*                                              set path to the current opened project, 
*                                              show the paths collected, and better logging
*                                              of the details. Added three methods to allow 
*                                              developers to configure extra pathing details
*                                              before the project is scanned.
* ----------------------------------------------------------------------------------------
* 06/20/2015  Richard A. Schummer     2.3      Default location for log file now in VFP 
*                                              temp folder. New log entries are included 
*                                              at beginning of the log file. Also, temp 
*                                              log for current pass is deleted to 
*                                              Recycle Bin
* ----------------------------------------------------------------------------------------
* 07/03/2015  Richard A. Schummer     3.0      Options dialog with generic settings and 
*                                              developer preferences:
*                                              display font name, size and attrib, log file 
*                                              name, whether you want logged, show the paths,
*                                              and root project folder added to path.
* ----------------------------------------------------------------------------------------
* 08/29/2015  Richard A. Schummer     3.1      Refactored performance. 
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     3.2      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS txParam1

* Used for Thor Registry Keys
#DEFINE ccTOOLNAME               "WLC Path Setter 3"

* Default settings for Thor Registry Keys
#DEFINE ccSETPATHERRLOGPATH      ADDBS(SYS(2023))
#DEFINE ccSETPATHERRLOGFILE      [WLCSetPathErrorLog.txt]
#DEFINE clSHOWPATHS              .F.
#DEFINE clLOGPATHS               .F.
#DEFINE cnSHOWPATHSFONTSIZE      8
#DEFINE ccSHOWPATHSFONTNAME      "Segoe UI"
#DEFINE ccSHOWPATHSFONTSTYLE     "N"
#DEFINE ccROOTPROJECTFOLDER      [J:\WLCProject]

* Extension class called at the end of the process
#DEFINE ccDEVHOOKPROGBASE        [RAS]
#DEFINE ccDEVELOPERHOOKPROG      [SetPathEnhancements"]  

* Generic settings used in the process
#DEFINE cnSETMEMOWIDTFORSHOWPATH 225
#DEFINE cnSECONDSINDAY           24*60*60
#DEFINE ccCRLF                   CHR(13) + CHR(10)

IF PCOUNT() = 1 AND 'O' = VARTYPE(txParam1) AND 'thorinfo' == LOWER(txParam1.Class)

   WITH txParam1
   
      * Required
      .Prompt          = 'Set Path to Project Paths v3' && used in menus
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2 && a description for the tool
         This program opens up a project and sets the appropriate pathing to the files in the project. This allows forms to be run standalone and compiles to take place without worry if the paths are set correctly.
         
         If a project is open, you are asked if you want the paths to be set up for this project. Otherwsie, if there is only one project file in the current folder, the project is automatically opened, otherwise the developer is prompted for the PJX they want to open. You can also set paths for the currently opened project.
      ENDTEXT  
      
      .StatusBarText   = "Build FoxPro paths from project paths"  
      .CanRunAtStartUp = .T.
      
      * Options Dialog settings
      .OptionTool      = ccTOOLNAME
      .OptionClasses   = "cusShowPaths, cusShowPathFontSize, cusShowPathFontName, cusShowPathFontStyle, cusLogFileName, cusRootProjectFolder"

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source          = "WLC"                   && where did this tool come from?  Your own initials, for instance
      .Category        = "WLC"                   && creates categorization of tools; defaults to .Source if empty
      .Sort            = 0                       && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version         = "Version 3.2 - September 20, 2015"                          && e.g., 'Version 7, May 18, 2011'
      .Author          = "Rick Schummer and Steve Sawyer"
      .Link            = "http://www.whitelightcomputing.com/resourcesdeveloper.htm" && link to a page for this tool
      .VideoLink       = ""                                                          && link to a video for this tool
   ENDWITH 

   RETURN txParam1
ENDIF 

IF PCOUNT() = 0
   DO ToolCode
ELSE
   DO ToolCode WITH txParam1
ENDIF

RETURN


********************************************************************************
*  METHOD NAME: ToolCode
*
*  AUTHOR: Richard A. Schummer, November 2014
*
*  METHOD DESCRIPTION:
*    Core processing to open the project and set the paths for the various
*    paths VFP keeps track of in the IDE.
*
*  INPUT PARAMETERS:
*    tcProjectName = character, really only useful for standalone run and not 
*                    running through Thor. For instance:
*
*                    DO ToolCode in Thor_Tool_SetPath.PRG WITH "sos.pjx"
*
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode(tcProjectName)

* Need these variables to survive the RELEASE ALL EXCEPT
PUBLIC gnProjectFilesAvailable, ;
       glOpenAlready, ;
       gcSetPathErrorLog

SET ASSERTS ON

* Make sure that the proper parameter(s) are passed
IF PCOUNT() = 1 AND VARTYPE(m.tcProjectName) = "C" AND NOT EMPTY(m.tcProjectName)
   ASSERT FILE(FORCEEXT(m.tcProjectName,"PJX")) MESSAGE "Can't find specified project"
   RETURN 
ELSE
   * RAS 17-Jun-2015, for Jody Meyer, added ability to set path for current project
   glOpenAlready = .F.
   llCanceled    = .F.

   TRY 
      loProject     = application.ActiveProject
      tcProjectName = LOWER(m.loProject.Name)

      lnResult = MESSAGEBOX("Do you want to set path for " + m.tcProjectName + "?", ; 
                            0+3+32+256, ;
                            _screen.Caption)
                 
      DO CASE
         CASE m.lnResult = 6
            * Yes
            glOpenAlready = .T.
            
         CASE m.lnResult = 7
            * No 
            * Continue on so the developer can select the project to open

         OTHERWISE
            * Cancel
            * Nothing to log at this time since we have not done anything.
            llCanceled = .T.

      ENDCASE           
   CATCH TO loException
      * Nothing to do as there is real possibility that a project is not open at this time.

   ENDTRY
   
   IF m.llCanceled = .T.
      RETURN 
   ENDIF 
   
   IF m.glOpenAlready
      * No more selection of the project, continue on...
   ELSE
      gnProjectFilesAvailable = ADIR(laPjx, "*.pjx")
   
      IF m.gnProjectFilesAvailable = 1
         tcProjectName = laPjx[1,1]
      ELSE
         tcProjectName = GETFILE("PJX", "Pick the project")
      ENDIF 
   ENDIF 
   
   IF EMPTY(tcProjectName)
      RETURN 
   ENDIF 
ENDIF

* Clean up and reset the environment
CLOSE ALL
RELEASE ALL EXCEPT tcProjectName
CLEAR
ON ERROR
ON SHUTDOWN
SET CLASSLIB TO
SET PROCEDURE TO
SET EXCLUSIVE OFF
SET DELETED ON
SET TALK OFF
SET ESCAPE ON

* Variable declarations.
LOCAL lcPath, ;
      loException AS Exception, ;
      lcProjDir, ;
      lcToolsDir, ;
      lcCurDir, ;  
      lcDeveloper, ;
      lcProjectSetPathAdditional, ;
      lcProjectFile, ;
      lcNewPath, ;
      lnFileCount, ;
      lnCounter, ;
      lnPercentage, ;
      lcClassLib, ;
      lcCode, ;
      lcDeveloperHookProgram, ;
      lcErrorLogInfo, ;
      lcFile, ;
      lcHomeDir, ;
      lcType, ;
      loFile, ;
      loProject, ;
      lnStartSeconds, ;
      lnEndSeconds, ;
      lnTimeToProcessProject, ;
      lnFiles, ;
      lcOldSafety

* Grab the developer from the login id.
lcDeveloper               = RIGHT(SYS(0), RAT(SPACE(1), SYS(0)) -1)
lcCurDir                  = FULLPATH(CURDIR())

* Use Thor to get some settings for generic folders and error files
lcProjDir                 = EXECSCRIPT(_Screen.cThorDispatcher, "Get Option=", "Root Project Folder", ccTOOLNAME)
lcProjDir                 = IIF(ISNULL(m.lcProjDir), SPACE(0), ALLTRIM(m.lcProjDir))
lcToolsDir                = IIF(ISNULL(m.lcProjDir) OR EMPTY(m.lcProjDir), SPACE(0), ADDBS(m.lcProjDir) + [Tools])
gcSetPathPermErrorLogFile = ALLTRIM(EXECSCRIPT(_Screen.cThorDispatcher, "Get Option=", "Log File Name", ccTOOLNAME))
gcSetPathPermErrorLogFile = IIF(ISNULL(m.gcSetPathPermErrorLogFile), SYS(2023), ALLTRIM(m.gcSetPathPermErrorLogFile))
gcSetPathPassErrorLogFile = FORCEEXT(JUSTPATH(m.gcSetPathPermErrorLogFile) + SYS(2015), "txt")

lnStartSeconds        = 0
lnEndSeconds          = 0

* If the permanent set path log file does not exist, initialize it to nothing.
IF FILE(gcSetPathPermErrorLogFile)
   * Nothing to do 
ELSE
   STRTOFILE(SPACE(0), gcSetPathPermErrorLogFile)
ENDIF 

* Start the path setting with non-project tool and add-on directories
* such as SDT, editors, browsers, etc. classes to the path.
* SAS - 09/28/01 - Changed to a full path so this will work
* when logged to a different directory other than main project folder

lcPath = SPACE(0)

IF NOT EMPTY(m.lcProjDir) AND DIRECTORY(m.lcProjDir) 
   lcPath = m.lcPath + IIF(EMPTY(m.lcPath), SPACE(0), [;]) + m.lcProjDir 
ENDIF 

IF NOT EMPTY(m.lcToolsDir) AND DIRECTORY(m.lcToolsDir) 
   lcPath = m.lcPath + IIF(EMPTY(m.lcPath), SPACE(0), [;]) + m.lcToolsDir
ENDIF 

* Add VFP home directory, version independent
lcPath   = m.lcPath + IIF(EMPTY(m.lcPath), SPACE(0), [;]) + HOME()

* First, add standard developer preference directories to the beginning of the path.
SetPathDeveloperPreferenceDirectories(m.lcPath)

* Add standard project directories to the path. This boosts performance for bigger projects.
SetPathCommonProjectDirectories()


* Get the path as it was built so far.
lcPath = SET("Path")

* SET CLASSLIB to any non-project classlibs here
AddCommonNonProjectClassLibs()

* SET PROCEDURE TO any non-project classlibs here
AddCommonSetProcedures()


* Assumption here is that you're working on a local drive, either
* in primary development or a source code controlled copy of the project
TRY
   llOK = .T.

   * Build some cursors to process folders faster
   TRY 
      USE (m.tcProjectName) IN 0 SHARED AGAIN ALIAS _SetPathProj
      
      * Build distinct list of folders in the project
      SELECT DISTINCT ;
             CAST(JUSTPATH(SYS(2014, LEFTC(name, 250))) AS C(250)) AS cFolder, ;
             Type AS cType ;
         FROM _setpathproj ;
         INTO CURSOR curFolders

      * Build list of files used for SET CLASSLIB and SET PROCEDURE (only class libraries and programs)
      SELECT CAST(FULLPATH(LEFTC(name, 250)) AS C(250)) AS cClassLib, ;
             Type AS cType ;
         FROM _setpathproj ;
         WHERE type = "V" ;
         INTO CURSOR curClassLibs
      
      SELECT CAST(FULLPATH(LEFTC(name, 250)) AS C(250)) AS cProgram, ;
             Type AS cType ;
         FROM _setpathproj ;
         WHERE type = "P" ;
         INTO CURSOR curPrograms

      USE IN (SELECT("_SetPathProj"))
   
   CATCH TO loException
      * Project is most likely already open
      lcCode = "Error: " + m.loException.Message + ;
               " [" + TRANSFORM(m.loException.Details) + "] " + ;
               " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
               " in " + m.loException.Procedure + ;
               " on " + TRANSFORM(m.loException.LineNo)

      MESSAGEBOX(lcCode, ;
                 0+48, ;
                 _screen.Caption)
    
      llOK = .F.
      
   ENDTRY

   TRY
      MODIFY PROJECT (m.tcProjectName) NOWAIT
      
   CATCH TO loException WHEN loException.ErrorNo = 1705
      lcCode = "Access is denied to the project file ( " + m.tcProjectName + "). It is probably open already in different session of Visual FoxPro."

      MESSAGEBOX(m.lcCode, ;
                 0+48, ;
                 _screen.Caption)
      
      llOK = .F.   
   
   CATCH TO loException 
      * Project is most likely already open
      lcCode = "Error: " + m.loException.Message + ;
               " [" + TRANSFORM(m.loException.Details) + "] " + ;
               " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
               " in " + m.loException.Procedure + ;
               " on " + TRANSFORM(m.loException.LineNo)

      MESSAGEBOX(lcCode, ;
                 0+48, ;
                 _screen.Caption)
    
      llOK = .F.
      
   ENDTRY 

   IF llOK
      loProject = application.ActiveProject
      lcHomeDir = UPPER(m.loProject.HomeDir)

      IF ADDBS(SET("Directory")) <> ADDBS(m.lcHomeDir)
         CD (m.lcHomeDir)
      ENDIF

      LogIssue("Opened " + ALLTRIM(m.tcProjectName) + " in project manager - " + TRANSFORM(DATETIME()), "L")

      * Extract all of the files, and use them to
      * determine the path, and SET PROCEDURE to all
      * PRG Files and SET CLASSLIB to all class libs
      IF FILE("liblist.dbf") && For storing program elements
         LOCAL lcOldSafety

         USE liblist IN 0 EXCLUSIVE
         lcOldSafety = SET("SAFETY")
         SET SAFETY OFF
         ZAP
         SET SAFETY &lcOldSafety
      ENDIF

      lnCounter   = 0
      lnStartSeconds = SECONDS()

      * SET PATH ------------------------------------------------------------
      SELECT curFolders
      
      WAIT WINDOW "Setting path from " + TRANSFORM(RECCOUNT("curClasslibs")) + " folders." NOWAIT 
      
      SCAN
         lcNewPath = ALLTRIM(curFolders.cFolder)
         IF NOT EMPTY(lcNewPath) AND (NOT (";" + lcNewPath) $ m.lcPath)
            lcPath = ALLTRIM(m.lcPath) + ";" + m.lcNewPath
         ENDIF
      ENDSCAN
      
      SET PATH TO &lcPath

      LogIssue("Finished setting path from " + TRANSFORM(RECCOUNT("curFolders")) + " folders.", "L")

      * SET CLASSLIB --------------------------------------------------------
      SELECT curClassLibs
      
      WAIT WINDOW "Setting classlib path from " + TRANSFORM(RECCOUNT("curClasslibs")) + " files."NOWAIT 
      
      SCAN
         * Make sure we have a SET PROCEDURE TO this file
         lcFile = ALLTRIM(curClassLibs.cClassLib)
         lcType = ALLTRIM(curClassLibs.cType)
         
         IF USED("liblist") 
            INSERT INTO liblist (mFileName, cType) VALUES (m.lcFile, m.lcType)
         ENDIF 
         
         IF NOT m.lcFile $ SET("ClassLib")
            TRY
               SET CLASSLIB TO (m.lcFile) ADDITIVE
            
            CATCH TO loException
               lcErrorLogInfo = TRANSFORM(DATETIME()) + ;
                                ": (form) " + ;
                                m.loException.Message + ;
                                " - " + ;
                                m.loException.Details + ;
                                ccCRLF  
               LogIssue(m.lcErrorLogInfo, "L")
            
            ENDTRY
         ENDIF
         
      ENDSCAN

      LogIssue("Finished setting classlib path from " + TRANSFORM(RECCOUNT("curClasslibs")) + " files.", "L")
      
      * SET PROCEDURE -------------------------------------------------------
      SELECT curPrograms
      
      SCAN
         * Make sure we have a SET PROCEDURE TO this file
         lcFile = ALLTRIM(curPrograms.cProgram)
         lcType = ALLTRIM(curPrograms.cType)
         
         IF USED("liblist") 
            INSERT INTO liblist (mFileName, cType) VALUES (m.lcFile, m.lcType)
         ENDIF 
         
         IF NOT m.lcFile $ SET("PROCEDURE")
            TRY
               SET PROCEDURE TO (m.lcFile) ADDITIVE
            
            CATCH TO loException
               lcErrorLogInfo = TRANSFORM(DATETIME()) + ;
                                ": (procedure) " + ;
                                m.loException.Message + ;
                                " - " + ;
                                m.loException.Details + ;
                                ccCRLF  
               LogIssue(m.lcErrorLogInfo, "L")
            
            ENDTRY
         ENDIF
      ENDSCAN
      
      LogIssue("Finished setting procedure list from " + TRANSFORM(RECCOUNT("curPrograms")) + " files.", "L")

   ENDIF 

   USE IN (SELECT("curFolders"))
   USE IN (SELECT("curClassLibs"))
   USE IN (SELECT("curPrograms"))

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
   lnEndSeconds = SECONDS()

ENDTRY

* RAS 17-Jun-2015, Added in the calculation of the time spent processing the project files, account for midnight project opening
IF lnStartSeconds = 0
   lnTimeToProcessProject = 0
ELSE 
   lnTimeToProcessProject = IIF(m.lnEndSeconds > m.lnStartSeconds, m.lnEndSeconds - m.lnStartSeconds, cnSECONDSINDAY - m.lnStartSeconds + m.lnEndSeconds)
ENDIF 

* RAS 10-Jan-2001 Added hook call to program so developers can customize their
*  additions like SDT and hotkeys in their very own customized program.
lcDeveloperHookProgram = FORCEEXT(ccDEVHOOKPROGBASE + ccDEVELOPERHOOKPROG, "prg")

IF FILE(m.lcDeveloperHookProgram)
   TRY
      DO (m.lcDeveloperHookProgram)

   CATCH TO loException
      lcErrorLogInfo = TRANSFORM(DATETIME()) + ;
                       ": (developerhook) " + ;
                       m.loException.Message + ;
                       " - " + ;
                       m.loException.Details + ;
                       ccCRLF
      LogIssue(m.lcErrorLogInfo, "L")

   ENDTRY
ELSE
   * No hooks created, just execute base startup code
ENDIF

* Reset the message 
SET MESSAGE TO 

IF ExecScript(_Screen. cThorDispatcher, "Get Option=", "Show Paths", ccTOOLNAME)
   ShowPaths()
ENDIF 

IF ExecScript(_Screen. cThorDispatcher, "Get Option=", "Log Paths", ccTOOLNAME)
   LogPaths()
ENDIF 

lcTimeMessage = "Processed set paths for " + LOWER(m.tcProjectName) + " in " + TRANSFORM(m.lnTimeToProcessProject) + " seconds..."
WAIT WINDOW (m.lcTimeMessage) NOWAIT 

lnFiles = application.ActiveProject.Files.Count
LogIssue(TRANSFORM(lnFiles) + " file" + IIF(lnFiles = 1, SPACE(0), "s") + " in the project", "L")
LogIssue(m.lcTimeMessage, "L")
LogIssue(REPLICATE("-", 80))

* RAS 20-Jun-2015, place the current log for this pass into the general log file along with the current contents.
lcOldSafety = SET("Safety")
SET SAFETY OFF 

STRTOFILE(FILETOSTR(gcSetPathPassErrorLogFile) + ccCRLF + FILETOSTR(m.gcSetPathPermErrorLogFile), ;
          m.gcSetPathPermErrorLogFile)

SET SAFETY &lcOldSafety

DELETE FILE (m.gcSetPathPassErrorLogFile) RECYCLE 

* Clean up the public variables.
RELEASE gnProjectFilesAvailable, gnOpenAlready, gcSetPathPassErrorLog, gcSetPathPermErrorLogFile
   
RETURN


********************************************************************************
*  METHOD NAME: SetPathDeveloperPreferenceDirectories
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    This method allows developer to add in their preferred paths to the 
*    beginning of the SET PATH. This is a perfect location to set paths up 
*    to their favorite tool folders.
*
*  INPUT PARAMETERS:
*    tcPath = character, not required, should be path to valid folders. No 
*             validation of the folders is done here.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE SetPathDeveloperPreferenceDirectories(tcPath)

* If something is passed in and a character, set it to the path. 
IF PCOUNT() = 1 AND VARTYPE(m.tcPath) = "C"
   SET PATH TO &tcPath 
ENDIF 

*? Developers can add more here

RETURN 


********************************************************************************
*  METHOD NAME: SetPathCommonProjectDirectories
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    Add native and standard project directories to the path first before checking
*    the project for folders. This will speed up the process a little bit since it
*    adds to the path once, not for each first file in the folder.
*
*    Initial design incorporates many common named subfolders from the project 
*    folder and only if they exist. The reason for not incorporating all 
*    subfolders is so folders for things like Deployment or backup data folders 
*    are not included automatically.
*
*    Initial testing shows this technique can save 33% of the time to open
*    larger projects and set the path to the project folders for the files
*    in the project.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE SetPathCommonProjectDirectories()

SET PATH TO "." ADDITIVE

IF DIRECTORY("programs")
   SET PATH TO "programs" ADDITIVE 
ENDIF  

IF DIRECTORY("prgs")
   SET PATH TO "prgs" ADDITIVE 
ENDIF  

IF DIRECTORY("progs")
   SET PATH TO "progs" ADDITIVE 
ENDIF  

IF DIRECTORY("data")
   SET PATH TO "data" ADDITIVE 
ENDIF  

IF DIRECTORY("libs")
   SET PATH TO "libs" ADDITIVE 
ENDIF  

IF DIRECTORY("classes")
   SET PATH TO "classes" ADDITIVE 
ENDIF  

IF DIRECTORY("vcx")
   SET PATH TO "vcx" ADDITIVE 
ENDIF  

IF DIRECTORY("forms")
   SET PATH TO "forms" ADDITIVE 
ENDIF  

IF DIRECTORY("reports")
   SET PATH TO "reports" ADDITIVE 
ENDIF  

IF DIRECTORY("frx")
   SET PATH TO "frx" ADDITIVE 
ENDIF  

IF DIRECTORY("labels")
   SET PATH TO "labels" ADDITIVE 
ENDIF  

IF DIRECTORY("lbx")
   SET PATH TO "lbx" ADDITIVE 
ENDIF  

IF DIRECTORY("graphics")
   SET PATH TO "graphics" ADDITIVE 
ENDIF  

IF DIRECTORY("images")
   SET PATH TO "images" ADDITIVE 
ENDIF  

IF DIRECTORY("bmp")
   SET PATH TO "bmp" ADDITIVE 
ENDIF  

IF DIRECTORY("menus")
   SET PATH TO "menus" ADDITIVE 
ENDIF  

IF DIRECTORY("includes")
   SET PATH TO "includes" ADDITIVE 
ENDIF  

IF DIRECTORY("H")
   SET PATH TO "H" ADDITIVE 
ENDIF  

IF DIRECTORY("text")
   SET PATH TO "text" ADDITIVE 
ENDIF  

IF DIRECTORY("test")
   SET PATH TO "test" ADDITIVE 
ENDIF  

IF DIRECTORY("help")
   SET PATH TO "help" ADDITIVE 
ENDIF  

IF DIRECTORY("metadata")
   SET PATH TO "metadata" ADDITIVE 
ENDIF  

IF DIRECTORY("docs")
   SET PATH TO "docs" ADDITIVE 
ENDIF  

RETURN 


********************************************************************************
*  METHOD NAME: 
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    SET CLASSLIB to any non-project classlibs here.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE AddCommonNonProjectClassLibs()

*< SET CLASSLIB TO "" ADDITIVE 

RETURN 


********************************************************************************
*  METHOD NAME: AddCommonSetProcedures
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    SET PROCEDURE TO any non-project classlibs here.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
* 
PROCEDURE AddCommonSetProcedures()

*< SET PROCEDURE TO "" ADDITIVE 

RETURN 


********************************************************************************
*  METHOD NAME: ShowPaths
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    Display the different set paths created by this program. This creates 
*    quite a dump of information for larger projects. Font and Fontsize for 
*    the text can be set using #DEFINEs at the top of the ToolCode procedure.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ShowPaths()

LOCAL lcOldScreenFontName, ;
      lcFontAttrib, ;
      llFontBold, ;
      llFontItalic, ;
      llOldScreenFontBold, ;
      llOldScreenFontItalic, ;
      lnOldMemoWidth, ;
      lnOldScreenFontSize

lnOldMemoWidth        = SET("Memowidth")
lnOldScreenFontSize   = _screen.FontSize
lcOldScreenFontName   = _screen.FontName
llOldScreenFontBold   = _screen.FontBold
llOldScreenFontItalic = _screen.FontItalic 

SET MEMOWIDTH TO cnSETMEMOWIDTFORSHOWPATH

lcFontAttrib = EXECSCRIPT(_Screen.cThorDispatcher, "Get Option=", "Show Path FontStyle", ccTOOLNAME)
lcFontAttrib = IIF(ISNULL(lcFontAttrib), "N", UPPER(lcFontAttrib))
llFontBold   = IIF("B" $ lcFontAttrib, .T., .F.)
llFontItalic = IIF("I" $ lcFontAttrib, .T., .F.)

_screen.FontName   = EXECSCRIPT(_Screen.cThorDispatcher, "Get Option=", "Show Path FontName", ccTOOLNAME)
_screen.FontSize   = EXECSCRIPT(_Screen.cThorDispatcher, "Get Option=", "Show Path FontSize", ccTOOLNAME)
_screen.FontBold   = llFontBold
_screen.FontItalic = llFontItalic

? "PATH:"
? SET("Path")
?
? "CLASS LIBRARIES:"
? SET("Classlib")
?
? "PROCEDURES:"
? SET("Procedure")
?

_screen.FontSize   = m.lnOldScreenFontSize
_screen.FontName   = m.lcOldScreenFontName
_screen.FontBold   = m.llOldScreenFontBold
_screen.FontItalic = m.llOldScreenFontItalic

SET MEMOWIDTH TO lnOldMemoWidth

RETURN 


********************************************************************************
*  METHOD NAME: LogPaths
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    Record the different set paths created by this program in the log file. 
*    This creates quite a dump of information for larger projects. 
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE LogPaths()

LOCAL lcLogText, ;
      lcSetPath, ;
      lcSetClassLib, ;
      lcSetProcedure

lcSetPath      = SET("Path")
lcSetClassLib  = SET("Classlib")
lcSetProcedure = SET("Procedure")

lcLogText = SPACE(0)

lcLogText = m.lcLogText + ;
            "PATH:" + ccCRLF + ;
            IIF(EMPTY(lcSetPath), "<empty>", lcSetPath) + ccCRLF + ccCRLF

lcLogText = m.lcLogText + ;
            "CLASS LIBRARIES:" + ccCRLF + ;
            IIF(EMPTY(lcSetClassLib), "<empty>", lcSetClassLib) + ccCRLF + ccCRLF

lcLogText = m.lcLogText + ;
            "PROCEDURES:" + ccCRLF + ;
            IIF(EMPTY(lcSetProcedure), "<empty>", lcSetProcedure) + ccCRLF 

LogIssue(m.lcLogText, "L")

RETURN 



********************************************************************************
*  METHOD NAME: LogIssue
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    This method is called to log any issues you want in the log file, or expressed
*    in a messagebox, or wait window, or dumped on the screen. All messages get
*    dumped into the Debug Output window if the debugger is open.
*
*  INPUT PARAMETERS:
*    tcMessage     = character, required, text to be recorded or displayed.
*    tcMessageType = character, (W)ait Window, (M)essageBox, (D)ebugout, (S)creen, 
*                    (L)og Filedefaults to "L" if not passed.
*    tlNoWait      = locical, not required, for wait window, to not introduce wait state
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE LogIssue(tcMessage, tcMessageType, tlNoWait)

LOCAL lcNoWait, ;
      lcDefaultMessageType
      
lcDefaultMessageType = "L"

ASSERT VARTYPE(m.tcMessage)="C" MESSAGE PROGRAM()+": Message is not character type"

* Handle setting default message type
IF VARTYPE(m.tcMessageType) = "C"
   * (W)ait Window, (M)essageBox, (D)ebugout, (S)creen, (L)og File
   IF INLIST(m.tcMessageType, "W", "M", "D", "S", "L")
   ELSE
      * Default to Log file
      tcMessageType = m.lcDefaultMessageType
   ENDIF
ELSE
   tcMessageType = m.lcDefaultMessageType
ENDIF

* Make sure the message can be displayed without wait state
IF m.tlNoWait
   lcNoWait = "NOWAIT"
ELSE
   lcNoWait = SPACE(0)
ENDIF

* Made it so the messages are always DEBUGOed
DEBUGOUT (m.tcMessage)

* Send the message to the requested location
DO CASE 
   CASE m.tcMessageType = "L"
      STRTOFILE(tcMessage + ccCRLF, gcSetPathPassErrorLogFile, 1)
      
   CASE m.tcMessageType = "W"
      WAIT WINDOW (tcMessage) &lcNoWait
   
   CASE m.tcMessageType = "M"
      MESSAGEBOX(tcMessage, 0+48, "SetPath Program")
   
   CASE m.tcMessageType = "S"
      ACTIVATE SCREEN 
      ? m.tcMessage
ENDCASE

RETURN


***********************************************************************************************************************
* Option classes: (notes from Mike Potjer and Jim Nelson)
*   - The Tool property corresponds to the OptionTool property.
*   - The Key property (combined with the Tool property) is what you will use to retrieve a specific setting from the
*     settings table.
*   - The Value property is the initial value for the setting. This does not have to be a string, but can be a data type
*     appropriate for the setting.
*   - The EditClassName property specifies the visual class displayed to the user for editing the setting.
*      + The edit class may, and often does, contain controls for more than one setting.
*      + The edit class does NOT have to be defined in the same .PRG as the option class. To use an edit class
*        defined in some other file, use this syntax:
*            EditClassName = ' clsMyToolOptionsPage of MySharedToolOptions.PRG'
*   - An option class MUST be defined in the same PRG as the tool that specifies it in its OptionClasses property. In
*     general, it does not work to subclass an option class from another file, because the class libraries used by Thor
*     are not typically in the path.
*   - Multiple tools may use the same setting (not just the tool that defined it), but they must all specify the same
*     Tool and Key combination when reading from or writing to that setting.
*   - It is not necessary to define an option class in every tool that needs access to a particular setting. However, at
*     least one option class must be defined for a tool if you want an “Options” link to appear on the tool’s page in the
*     Tool Launcher.
***********************************************************************************************************************
DEFINE CLASS cusSetPathBase AS Custom
   Tool          = ccTOOLNAME
   EditClassName = "ctrSetPath3OptionsPage" 
ENDDEFINE

DEFINE CLASS cusShowPaths AS cusSetPathBase
   Key   = "Show Paths"
   Value = clSHOWPATHS
ENDDEFINE

DEFINE CLASS cusShowPathFontSize AS cusSetPathBase
   Key   = "Show Path FontSize"
   Value = cnSHOWPATHSFONTSIZE
ENDDEFINE

DEFINE CLASS cusShowPathFontName AS cusSetPathBase
   Key   = "Show Path FontName"
   Value = ccSHOWPATHSFONTNAME
ENDDEFINE

DEFINE CLASS cusShowPathFontStyle AS cusSetPathBase
   Key   = "Show Path FontStyle"
   Value = ccSHOWPATHSFONTSTYLE
ENDDEFINE

DEFINE CLASS cusLogFileName AS cusSetPathBase
   Key   = "Log File Name"
   Value = ccSETPATHERRLOGPATH + ccSETPATHERRLOGFILE
ENDDEFINE

DEFINE CLASS cusRootProjectFolder AS cusSetPathBase
   Key   = "Root Project Folder"
   Value = ccROOTPROJECTFOLDER
ENDDEFINE


*********************************************************************
*-- Define the option page for all the SetPath options.
*********************************************************************
DEFINE CLASS ctrSetPath3OptionsPage AS Container

PROCEDURE Init()
   LOCAL loRenderEngine

   * Get a reference to the Thor OptionRenderEngine (Dynamic Forms)
   loRenderEngine = EXECSCRIPT( _Screen.cThorDispatcher, "Class= OptionRenderEngine" )

   
   * Create the Dynamic Forms Options dialog markup syntax
   TEXT TO loRenderEngine.cBodyMarkup NOSHOW TEXTMERGE
      .Class         = "CheckBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .BackStyle     = 0
      .Caption       = "Display paths? (dumped on screen)"
      .Width         = 250
      .cTool         = ccToolName
      .cKey          = "Show Paths"
      |

      .Class         = "Label"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Caption       = "Show path font size and name:"
      .WordWrap      = .F.
      .Width         = 250
      |

      .Class         = "Spinner"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .InputMask     = "99"
      .BackStyle     = 0
      .Name          = "spnFontSize"
      .Width         = 50
      .cTool         = ccToolName
      .cKey          = "Show Path FontSize"
      |

      .Class         = "TextBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Name          = "txtFontName"
      .Width         = 175
      .cTool         = ccToolName
      .cKey          = "Show Path FontName"
      .row-increment = '0'     
      |

      .Class         = "TextBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Name          = "txtFontStyle"
      .Width         = 30
      .cTool         = ccToolName
      .cKey          = "Show Path FontStyle"
      .row-increment = '0'     
      |

      .Class         = "cmdGetFontName"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .row-increment = '0'     
      |

      .Class         = "CheckBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .BackStyle     = 0
      .Caption       = "Log the paths? (details recorded in log file)"
      .Width         = 250
      .cTool         = ccToolName
      .cKey          = "Log Paths"
      |

      .Class         = "lblLogFile"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Caption       = "Log File: "
      .WordWrap      = .F.
      .Width         = 250
      |

      .Class         = "TextBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Name          = "txtLogPath"
      .Width         = 300
      .cTool         = ccToolName
      .cKey          = "Log File Name"
      |

      .Class         = "cmdLogFilePicker"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .row-increment = '0'     
      |

      .Class         = "Label"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Caption       = "Root VFP Project Folder (added to beginning of path): "
      .WordWrap      = .F.
      .Width         = 300
      |

      .Class         = "TextBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Name          = "txtProjectFolder"
      .Width         = 300
      .cTool         = ccToolName
      .cKey          = "Root Project Folder"
      |

      .Class         = "cmdGetProjectFolder"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .row-increment = '0'     
      |

   ENDTEXT

   * Render this container.
   loRenderEngine.Render(This, ccToolName )
ENDPROC

ENDDEFINE


*************************************************************************************
DEFINE CLASS cmdLogFilePicker AS CommandButton

Caption  = "..."
Width    = 24

PROCEDURE Click()

LOCAL lcOldDir, ;
      lcFileName

TRY 
   lcOldDir   = FULLPATH(CURDIR())
   
   CD (JUSTPATH(this.Parent.txtLogPath.Value))
   
   lcFileName = GETFILE("txt", "Log File")
   
   CD (m.lcOldDir)
   
   IF EMPTY(m.lcFileName)
      * Nothing to do
   ELSE
      lcLogFile = LOWER(ALLTRIM(m.lcFileName))
      
      * Set to object value for physical presentation, then save to the registry to be permanent.
      this.Parent.txtLogPath.Value = m.lcLogFile

      ExecScript(_Screen.cThorDispatcher, "Set Option=", "Log File Name", ccTOOLNAME, m.lcLogFile)
   ENDIF 
   
CATCH TO loException
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(lcCode, ;
              0+48, ;
              _screen.Caption)
 
ENDTRY

ENDPROC

ENDDEFINE


*************************************************************************************
DEFINE CLASS cmdGetProjectFolder AS CommandButton

Caption  = "..."
Width    = 24

PROCEDURE Click()

LOCAL lcFolderName, ;
      lcProjectFolder

TRY 
   lcFolderName = GETDIR(ALLTRIM(this.Parent.txtProjectFolder.Value), "Root Project Folder")
   
   IF EMPTY(m.lcFolderName)
      * Nothing to do
   ELSE
      lcProjectFolder = LOWER(ALLTRIM(m.lcFolderName))
      
      * Set to object value for physical presentation, then save to the registry to be permanent.
      this.Parent.txtProjectFolder.Value = m.lcProjectFolder

      ExecScript(_Screen.cThorDispatcher, "Set Option=", "Root Project Folder", ccTOOLNAME, m.lcProjectFolder)
   ENDIF 
   
CATCH TO loException
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(lcCode, ;
              0+48, ;
              _screen.Caption)

ENDTRY

RETURN 

ENDPROC

ENDDEFINE 


*************************************************************************************
DEFINE CLASS cmdGetFontName AS CommandButton

Caption  = "Font..."
Width    = 45

PROCEDURE Click()

LOCAL lcFontAttrib, ;
      lcFontName, ;
      lnFontSize

TRY 
   * Pass in the existing font, Returns Segoe UI,12,B
   lcFontAttrib = GETFONT(this.Parent.txtFontName.Value, this.Parent.spnFontSize.Value, this.Parent.txtFontStyle.Value)
   
   IF EMPTY(lcFontAttrib)
      * Nothing to do
   ELSE
      * Just return the font name, first, then set the font size
      lcFontName  = GETWORDNUM(m.lcFontAttrib, 1, ",")
      lnFontSize  = VAL(GETWORDNUM(m.lcFontAttrib, 2, ","))
      lcFontStyle = GETWORDNUM(m.lcFontAttrib, 3, ",")
      
      * Set to object value for physical presentation, then save to the registry to be permanent.
      this.Parent.txtFontName.Value  = m.lcFontName
      this.Parent.spnFontSize.Value  = m.lnFontSize
      this.Parent.txtFontStyle.Value = m.lcFontStyle
      
      ExecScript(_screen.cThorDispatcher, "Set Option=", "Show Path FontName", ccTOOLNAME, m.lcFontName)
      ExecScript(_screen.cThorDispatcher, "Set Option=", "Show Path FontSize", ccTOOLNAME, m.lnFontSize)
      ExecScript(_screen.cThorDispatcher, "Set Option=", "Show Path FontStyle", ccTOOLNAME, m.lcFontStyle)
   ENDIF 
   
CATCH TO loException
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(lcCode, ;
              0+48, ;
              _screen.Caption)

ENDTRY

RETURN 

ENDPROC

ENDDEFINE


*************************************************************************************
DEFINE CLASS lblLogFile AS Label

PROCEDURE Init()

this.ForeColor     = RGB(0,0,255)
this.MousePointer  = 15
this.FontUnderline = .T.
this.AutoSize      = .T.
   
ENDPROC 

PROCEDURE Click()
TRY 
   MODIFY FILE (this.Parent.txtLogPath.Value) NOEDIT RANGE 0,0 

CATCH TO loException   
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(lcCode, ;
              0+48, ;
              _screen.Caption)
 
ENDTRY

RETURN 

ENDPROC

ENDDEFINE



*: EOF :*   