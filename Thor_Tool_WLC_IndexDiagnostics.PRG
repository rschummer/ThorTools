******************************************************************************************
*  PROGRAM: Thor_Tool_WLC_IndexDiagnostics.prg
*
*  AUTHOR: Richard A. Schummer, January 2015
*
*  COPYRIGHT © 2015   All Rights Reserved.
*     White Light Computing, Inc.
*     Rick Schummer
*     PO Box 391
*     Washington Twp., MI  48094
*
*     raschummer@whitelightcomputing.com
*
*  COPYRIGHT © 2015   All Rights Reserved.
*     David G. Hanna
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
*     This Thor tool is designed to check the health of the indexes for tables. There
*     are three different ways this program runs, in this priority of determining how
*     the tables for analysis are determined.
*
*        1) DBCX Metadata: if CoreMeta.DBF is found, the table list is determined from 
*                          the tables registered in the DBCX/SDT metadata
*
*        2) DBC:           if a single DBC is found in the folder, then the DBC is used.
*                          If multiple DBCs are found, then the Windowss File Open dialog
*                          is opened to pick a database container.
*
*        3) Folder:        if no DBCX metadata files are found, and no DBC files are found, 
*                          the table list is determined from teh DBFs in the selected folder.
*
*     Logs potential issues with tags for things like DTOC(), ALLTRIM(), RTRIM() and STR()
*
*     Does not check for potential issues with NULL capable columns indexed, or UDFs.
*
*     More details about how the diagnostics work are found in the InspectIndex method
*     comment header.
*
*     SPECIAL THANKS to David Hanna for the core index diagnostics code for the two passes.
*     With his permission we extended it a little bit to help find other issues that trip
*     developers and even tools like Stonefield Database Toolkit (older versions).
*
*
*  CALLING SYNTAX:
*     This program works fine outside of Thor, but was designed with Thor integration
*
*     DO Thor_Tool_WLC_IndexDiagnostics.prg
*     Thor_Tool_WLC_IndexDiagnostics()
*
*  INPUT PARAMETERS:
*     lxParam1 = unknown type, not required, standard Thor parameter used to help 
*                configure Thor's tool menu.
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
*                                 C H A N G E    L O G                                    
*
*    Date     Developer               Version  Description
* ----------  ----------------------  -------  ---------------------------------
* 01/31/2015  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 02/07/2015  Richard A. Schummer     1.1      Clean up of the code and more testing
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     2.2      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

#DEFINE ccCRLF            CHR(13)+CHR(10)
#DEFINE ccTAB             CHR(9)
#DEFINE ccMsgCaption      "Index Integrity Checker"
#DEFINE ccLOGFILE         "CDXIntegrityCheckLog.txt"
#DEFINE ccALTERNATEFILE   "CDXIntegritySingleFile.txt"
#DEFINE cnMAXFAILS        10
#DEFINE cnTWOGIGLIMIT     (2*1024*1024*1024) -1
#DEFINE cnHOWCLOSETOLIMIT .50

#DEFINE ccDBCXSOURCE      "DBCX Metadata"
#DEFINE ccDBCSOURCE       "Database Container"
#DEFINE ccFOLDERSOURCE    "Windows Folder"

#DEFINE clTABDELIMITED    .F.

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND 'O' = VARTYPE(lxParam1) AND 'thorinfo' == LOWER(lxParam1.Class)
   WITH lxParam1
   
      * Required
      .Prompt          = 'Check Index Integrity'           && used in menus
      .StatusBarText   = 'Evaluate the indexes in database container, folder, or using DBCX metadata.'  
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2     && a description for the tool
         This program evaluates the indexes in database container, folder, or using DBCX metadata.      
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source        = "WLC"                      && where did this tool come from?  Your own initials, for instance
      .Category      = "WLC"                      && creates categorization of tools; defaults to .Source if empty
      .Sort          = 0                          && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version       = "Version 1.2, September 20, 2015"          && e.g., 'Version 7, May 18, 2011'
      .Author        = "Rick Schummer"
      .Link          = "https://github.com/rschummer/ThorTools"   && link to a page for this tool
      .VideoLink     = SPACE(0)                                   && link to a video for this tool
      
   ENDWITH 

   RETURN lxParam1
ENDIF 

IF PCOUNT() = 0
   DO ToolCode
ELSE
   DO ToolCode WITH lxParam1
ENDIF

RETURN


********************************************************************************
*  METHOD NAME: ToolCode
*
*  AUTHOR: Richard A. Schummer,
*
*  METHOD DESCRIPTION:
*    This is a wrapper program that loops through all the tables in a DBC to 
*    check if there is index issues in the different tables.
*
*  INPUT PARAMETERS:
*    lxParam1 = unknown type, optional, standard parameter passed in by Thor.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode(lxParam1)

LOCAL lcLogText, ;
      lcCode, ;
      lcDBCFile, ;
      lcTableToCheck, ;
      lnDBCFiles, ;
      lnJ, ;
      lnTables, ;
      loException AS Exception, ;
      lcOldSafety, ;
      lcMaxNumberOfFails, ;
      lnMaxNumberOfFails

TRY
   USE IN (SELECT("curPass1"))
   USE IN (SELECT("curPass2"))
   USE IN (SELECT("curGeneral"))
   USE IN (SELECT("CoreMeta"))

   * Most of the time I use this for the current folder, but allow developer option to pick
   lcFolderForFiles = GETDIR()
   
   IF EMPTY(lcFolderForFiles)
      RETURN 
   ENDIF 


   *  Create a cursor called "curPass1" to hold the Pass 1 results of the test
   CREATE CURSOR curPass1 ;
      ( cTable       C(200), ;
        nRecNo       N(8), ;
        cTagName     C(10), ;
        cTagType     C(10), ;
        cIdxVal      C(15), ;
        cMessage     C(25), ;
        cDupKey      C(1), ;
        cForErr      C(1) )

   *  Create a cursor called "curPass2" to hold the Pass 2 results of the test
   CREATE CURSOR curPass2 ;
      ( cTable       C(200), ;
        nRecNo       N(8), ;
        cTagName     C(10), ;
        cTagType     C(10), ;
        cIdxVal      C(15), ;
        cMessage     C(50) )

   * Create a cursor called curGeneral for some general statistics about each table analyzed.
   CREATE CURSOR curGeneral ;
      ( cTable       C(200), ;
        nRecord      I, ;
        nDeleted     I, ;
        nSeekErr     I, ;
        nForErr      I, ;
        nDupKey      I, ;
        nScanErr     I, ;
        nFoxErr      I, ;
        nElapsedTm   N(12,4) )

   lcLogText  = SPACE(0)
   lnDBCFiles = ADIR(laDatabases, "*.DBC")

   * Determine the source of the table list.
   DO CASE
      * ===============================================================
      * First, check if DBCX files and use list from CoreMeta
      CASE FILE("CoreMeta.DBF")
         SELECT * ;
            FROM coremeta ;
            WHERE cRecType = "T" ;
            ORDER BY cDBCName, cObjectNam ;
            INTO CURSOR curTables

         SELECT CAST(IIF(EMPTY(cDBCName), cObjectNam, ALLTRIM(cDBCName) + "!" + cObjectNam) AS C(200)) AS cTablePath ;
            FROM curTables ;
            INTO ARRAY laTables

         lnTables          = RECCOUNT("curTables")
         lcTableSource     = ccDBCXSOURCE
         lcSourceSpecifics = "DBCX location: " + JUSTPATH(FULLPATH(DBF("coremeta")))
         
         USE IN (SELECT("coremeta"))

      * ===============================================================
      * Second, check if DBC files (which means you lose free tables)
      CASE lnDBCFiles > 0
         IF lnDBCFiles = 1
            * If only one database in the folder, go with it.
            lcDBCFile = laDatabases[1, 1]
         ELSE
            * Otherwise let the user select.
            lcDBCFile = GETFILE("dbc")
         ENDIF 
       
         IF EMPTY(lcDBCFile)
            * Nothing to do
            lnTables = 0
         ELSE
            WAIT WINDOW "Opening database, please wait..." NOWAIT 
            
            OPEN DATABASE (lcDBCFile) SHARED 
            SET DATABASE TO (JUSTSTEM(lcDBCFile))

            WAIT WINDOW "Determining number of tables to analyze from database" NOWAIT            
            lnTables = ADBOBJECTS(laTables, "TABLE")
         ENDIF 

         lcTableSource     = ccDBCSOURCE
         lcSourceSpecifics = "DBC: " + FULLPATH(DBC())
         
      * ===============================================================
      * Lastly, get list of all DBFs in folder and work with that list
      OTHERWISE
         * RAS 01-Feb-2015, Already picked folder above
         *< lcFolderForFiles = GETDIR()
         
         *< IF EMPTY(lcFolderForFiles)
         *<    lnFilesInFolder = 0
         *< ELSE 

            lnFilesInFolder = ADIR(laFiles, "*.DBF")
            
            IF lnFilesInFolder = 0
               lnTables = 0               
            ELSE
               DIMENSION laTables[ALEN(laFiles, 1)]
               
               FOR lnI = 1 TO lnFilesInFolder
                  laTables[lnI] = laFiles[lnI, 1]
               ENDFOR 
               
               lnTables = ALEN(laTables, 1)
            ENDIF 
         *< ENDIF 
         
         lcTableSource     = ccFOLDERSOURCE
         lcSourceSpecifics = "Folder selected: " + LOWER(lcFolderForFiles)

   ENDCASE

   * Get from the user the number of maximum fails you want to record for any one table.
   lcDefaultNumberOfFails = TRANSFORM(cnMAXFAILS)
   lcMaxNumberOfFails     = INPUTBOX("Please enter in the maximum number of index seek fails:", ;
                                     ccMsgCaption, lcDefaultNumberOfFails, 20000, ;
                                     lcDefaultNumberOfFails, lcDefaultNumberOfFails)

   TRY 
      lnMaxNumberOfFails = VAL(lcMaxNumberOfFails)
   
   CATCH TO loException
      lnMaxNumberOfFails = VAL(cnMAXFAILS)

   ENDTRY

   * Sort the list and initialize the beginning of the log text
   IF lnTables > 0
      ASORT(laTables)
   ENDIF 

   lcLogText = lcLogText + ;
               ccMsgCaption + ccCRLF + ; 
               "Started: " + TRANSFORM(DATETIME()) + ccCRLF + ;
               "Source: " + lcTableSource + ccCRLF + ; 
               "Current Folder: " + FULLPATH(CURDIR()) + ccCRLF + ;
               "Number of tables to assess: " + TRANSFORM(lnTables)  + ccCRLF + ;
               "Maximum number of SEEK fails: " + TRANSFORM(lnMaxNumberOfFails) + ccCRLF + ;
               "Path: " + LOWER(SET("Path")) + ccCRLF + ccCRLF + ;
               lcSourceSpecifics + ccCRLF + ccCRLF

   lcLogText = lcLogText + ;
               "**** Search for 'ERROR:', 'WARNING:'  for any issues ****" + ccCRLF + ccCRLF + ;
               "You can review curPass1, curPass2, curGeneral cursors for additional analysis...." + ccCRLF + ;
               "This routine does not check for potential NULL value tags, and does not check any UDF calls with variable outcomes." + ;
               ccCRLF + ccCRLF ;
   
   ************************************************************************************
   * Loop for all the tables.
   FOR lnJ = 1 TO lnTables
      * Determine the table name based on the type of source file list
      DO CASE
         * =================================
         CASE lcTableSource = ccDBCXSOURCE
            IF "!" $ laTables[lnJ]
               laTables[lnJ] = alltrim(laTables[lnJ])
               
               * DBC contained table
               lcDBCName   = ALLTRIM(SUBSTRC(laTables[lnJ], 1, ATC("!", laTables[lnJ]) -1 ))
               lcTableName = ALLTRIM(SUBSTRC(laTables[lnJ], ATC("!", laTables[lnJ]) +1))
               
               TRY 
                  OPEN DATABASE (lcDBCName) SHARED 
                  
               CATCH TO loException               

               ENDTRY
               
               TRY 
                  lcTableToCheck = DBGETPROP(lcTableName, "TABLE","Path")
               
               CATCH TO loException
                  lcTableToCheck = lcTableName
                  
               ENDTRY

            ELSE
               * Free Table
               lcTableToCheck = laTables[lnJ]
            ENDIF 

         * =================================
         CASE lcTableSource = ccDBCSOURCE
            * DBC is already open 
            lcTableToCheck = DBGETPROP(laTables[lnJ], "TABLE","Path")
            

         * =================================
         OTHERWISE
            * Folder full of DBFs
            lcTableToCheck = laTables[lnJ]
      ENDCASE
   

      * Initialize the log for a specific table
      lcLogText      = lcLogText + ccCRLF + ;
                       PADL(TRANSFORM(lnJ), LENC(TRANSFORM(lnTables)), "0") + ") " + ;
                       REPLICATE("-", 70) + ccCRLF + ccCRLF + ;
                       "Checking indexes for table: " + UPPER(m.lcTableToCheck) + ccCRLF 
      
      
      * Run the inspection of the indexes for a table
      TRY 
         lnOldMemoWidth = SET("Memowidth")
         lcOldConsole   = SET("Console")
         
         SET MEMOWIDTH TO 225
         SET CONSOLE OFF 
         
         DO InspectIndex WITH ALLTRIM(m.lcTableToCheck), .T., m.lnMaxNumberOfFails
     
         SET CONSOLE &lcOldConsole
         SET MEMOWIDTH TO (lnOldMemoWidth)
         
      CATCH TO loException      
         lcCode = "ERROR: " + m.loException.Message + ;
                  " [" + TRANSFORM(m.loException.Details) + "] " + ;
                  " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
                  " in " + m.loException.Procedure + ;
                  " on " + TRANSFORM(m.loException.LineNo)
         
         lcLogText = lcLogText + ccCRLF + lcCode + ccCRLF 
         
         * In case the log file in InspectTable was not closed via error
         SET ALTERNATE OFF 
         SET ALTERNATE TO 
         
         * Error might leave the table open so attempt to close it.
         USE IN (SELECT(m.lcTableToCheck))
         
      ENDTRY

      * Add in the text from the index alternate file log.
      TRY 
         lcLogText = lcLogText + ccCRLF + ;
                     FILETOSTR(ccALTERNATEFILE) + ;
                     ccCRLF 
      
      CATCH TO loException      
         lcCode = "ERROR: " + m.loException.Message + ;
                  " [" + TRANSFORM(m.loException.Details) + "] " + ;
                  " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
                  " in " + m.loException.Procedure + ;
                  " on " + TRANSFORM(m.loException.LineNo)
         
         lcLogText = lcLogText + ccCRLF + lcCode + ccCRLF 
                     
      ENDTRY
   ENDFOR 
   
   WAIT CLEAR 

   * Summary information
   lcLogText = lcLogText + ;
               ccCRLF + ccCRLF + ;
               REPLICATE("-", 80) + ccCRLF + ccCRLF + ;
               TRANSFORM(lnTables) + " tables analyzed" + ccCRLF + ;
               ccMsgCaption + " is complete: " + TRANSFORM(DATETIME()) + ccCRLF + ccCRLF + ;
               "*: EOF :*"

   CLEAR 
   
   * Write the log file
   lcOldSafety = SET("Safety")
   SET SAFETY OFF 
   
   STRTOFILE(lcLogText, ccLOGFILE, 0)
   
   SET SAFETY &lcOldSafety
   
   * Review results in log file
   MODIFY FILE (ccLOGFILE) NOEDIT NOWAIT RANGE 1,1
   
   * Position analysis cursors
   GO TOP IN curPass1
   GO TOP IN curPass2
   
   * Close the cursor with the list of tables if it exists.
   USE IN (SELECT("curTables"))

   
CATCH TO m.loException
   lcCode = "Error: " + m.loException.Message + ;
      " [" + TRANSFORM(m.loException.Details) + "] " + ;
      " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
      " in " + m.loException.Procedure + ;
      " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(m.lcCode, ;
              0+48, ;
              _screen.Caption)

ENDTRY 

RETURN 


********************************************************************************
*  METHOD NAME: InspectIndex
*
*  AUTHOR: David Hanna
*
*  METHOD DESCRIPTION:
*     This program opens the specified table and tests its compound structural
*     index (.CDX) for index corruption. It is intended to be used via the
*     FoxPro IDE, but it could be incorporated into an executable.
*
*     The program executes two passes.
*
*     Pass 1. Using each defined index and each physical record, determine
*     the corresponding SEEK value and SEEK the record. Record any SEEK
*     failures that were detected in a cursor called "curPass1".
*
*     Pass 2. Using each defined index, SCAN through the table to ensure
*     each record is selected once and only once. Record the results in a
*     cursor called "curPass2".
*
*     As the program is running, it echos summary information to the screen.
*
*     When the program terminates it closes the table it opened and leaves
*     open the cursors "curPass1" and "curPass2". It also creates and
*     populates the following PUBLIC variables for the last table analysis:
*
*        PUBLIC gnRecordCnt        && Number of records in table
*        PUBLIC gnDeletedCnt       && Number of deleted records in table
*        PUBLIC gnSeekErrCnt       && Number of Pass 1 SEEK errors
*        PUBLIC gnForErrCnt        && Number of Pass 1 index filter errors
*        PUBLIC gnDupKeyCnt        && Number of Pass 1 duplicate keys
*        PUBLIC gnScanErrCnt       && Number of Pass 2 SCAN errors
*        PUBLIC gnFoxErrCnt        && Number of FoxPro errors trapped
*        PUBLIC gnElapsedTime      && Elapsed time for the entire test (seconds)
*
*     These values are stored in the curGeneral cursor for each table.
*
*  Required Parameters:
*
*     tcTablePath         * Input  Full pathname to table. The ".dbf"
*                                  extension can be provided, but is
*                                  not required.
*
*
*  Optional Parameters:
*
*     tlNoPrompt          * Input  If this parameter is not provided or is
*                                  .F., the program will display a prompt
*                                  when it competes the tests asking the
*                                  user for whether to terminate or display
*                                  a result cursor.
*                                  If this parameter is .T., the prompt is
*                                  not displayed and the program simply
*                                  terminates when it completes the tests.
*
*     tnMaxErrCnt          * Input This parameter is used to limit the
*                                  number of Pass 1 errors that are recorded
*                                  for any single index. If this parameter is
*                                  not provided, is not numeric, or is not
*                                  greater than 0, a default value of 10 will
*                                  be used. A special value of -1 can be used
*                                  to specify "no limit".
*                                  
*
*  Function Return:  -N-  Error code.
*                         0 = No errors were detected.
*                        -1 = One or more errors were detected.
*
*
*  Sample Usage:
*     DO InspectIndex WITH "c:\temp\MyTable.dbf"            && Interactive
*     DO InspectIndex WITH "c:\temp\MyTable.dbf", .T.       && Programmatic
*     DO InspectIndex WITH "c:\temp\MyTable.dbf", .T., 2    && Programmatic, max 2 errors
*
*
*  FoxPro Version:
*     This program can be run in FoxPro 2.6 or higher.
*
*     Note: A few variable names are longer than 10 characters, but the
*     first 10 characters are always unique.
*
*
*  Restrictions:
*     Tables with unique or candididate indexes are not fully supported.
*
*
*  Date         Revisor       Change
*  -----------  ------------  -----------------------------------------------
*  2008-Aug-20  D. Hanna      Created.
*  2008-Dec-27  D. Hanna      For Pass 1, only add errors to results cursor.
*  2009-Jul-08  D. Hanna      Show elapsed time, format displayed numbers .
*  2009-Aug-23  D. Hanna      Add ON ERROR to trap FoxPro SEEK() errors.
*  2009-Sep-30  D. Hanna      Warn user if any index expression uses DTOC().
*  2009-Oct-07  D. Hanna      Add "gnRecordCnt". Add check for maximum number
*                             of errors.
*  2010-Dec-07  D. Hanna      Fix bug where the work area for the table being
*                             checked was not selected before calling
*                             DESCENDING() and UNIQUE().
*  2015-Jan-31  Rick Schummer Modified to incorporate into a Thor tool.
*  2015-Feb-1   Rick Schummer Modified to handle generic logging via Alternate
*  2015-Feb-4   Rick Schummer MOdified to check for Primary/Candidate duplicates
*
*
******************************************************************************
PROCEDURE InspectIndex(tcTablePath, tlNoPrompt, tnMaxErrCnt)

LOCAL loException as Exception, ;
      lcOldSafety, ;
      llPrimary, ; 
      llCandidate

lcOldSafety = SET("Safety")
SET SAFETY OFF

SET ALTERNATE TO ccALTERNATEFILE
SET ALTERNATE ON 

*
*  Declare public variables
*  ------------------------
PUBLIC gnRecordCnt        && Number of records in table
PUBLIC gnDeletedCnt       && Number of deleted records in table
PUBLIC gnSeekErrCnt       && Number of Pass 1 SEEK errors
PUBLIC gnForErrCnt        && Number of Pass 1 index filter errors
PUBLIC gnDupKeyCnt        && Number of Pass 1 duplicate keys
PUBLIC gnScanErrCnt       && Number of Pass 2 SCAN errors
PUBLIC gnFoxErrCnt        && Number of FoxPro errors trapped
PUBLIC gnElapsedTime      && Elapsed time for the entire test (seconds)


*
*  Declare private variables
*  -------------------------
PRIVATE llNoPrompt        && Flag indicating whether to prompt user at end
PRIVATE lnMaxErrCnt       && Maximum errors allowed for any single index
PRIVATE lcOldTalk         && Old value of SET( "TALK" )
PRIVATE lcOldEsc          && Old value of SET( "ESCAPE" ) 
PRIVATE lcOldDel          && Old value of SET( "DELETED" )
PRIVATE lcOldNear         && Old value of SET( "NEAR" )
PRIVATE lcOldError        && Old value of ON( "ERROR" )
PRIVATE lnErrorNo         && FoxPro error number
PRIVATE lcAlias           && Alias of input table
PRIVATE lnSecStart        && Starting time of test (seconds)
PRIVATE lnRecCnt          && Number of records in input table
PRIVATE lnTagCnt          && Number of index tags in input table
PRIVATE lcTagName         && Current index tag name
PRIVATE lcIndexType       && Current index type ("Regular", or "Unique")
PRIVATE lcIndexExpr       && Current index expression
PRIVATE lcForExpr         && Current index filter expression
PRIVATE llDescend         && Flag indicating whether the index is DESCENDING
PRIVATE llUnique          && Flag indicating whether the index is UNIQUE
PRIVATE lnRecNo           && Current record number
PRIVATE luSeekKey         && Current SEEK key value (actual value)
PRIVATE lcSeekKey         && Current SEEK key value (as a string)
PRIVATE llForValue        && Current index filter value (T/F)
PRIVATE llFound           && Flag indicating the results of a SEEK()
PRIVATE llError           && Flag indicating whether an error was detected
PRIVATE lcMessage         && Message coresponding to a given error
PRIVATE lcForErr          && Flag indicating a index filter error was detected
PRIVATE lcDupKey          && Flag indicating a duplicate SEEK key was detected
PRIVATE lnKey             && Value of keyboard key pressed by user
PRIVATE lnP2LastRec       && Previous record number during PASS 2 SCAN
PRIVATE luP2LastKey       && Previous record SEEK key during PASS 2 SCAN (actual value)
PRIVATE lnP2RecCnt        && Number of records found in Pass 2
PRIVATE lnLoopCnt         && General purpose loop counter
PRIVATE llCanceled        && Flag indicating whether user canceled the test
PRIVATE lnIdxErrCnt       && Number of errors for current index
PRIVATE llFoxError        && Flag indicating a FoxPro error was trapped
PRIVATE llStopTest        && Flag indicating the current test must be stopped
PRIVATE llOk              && Flag indicating whether is is okay to proceed
PRIVATE i                 && General purpose loop index


*
*  Check the "no prompt" and "max errors" parameters
*  -------------------------------------------------
m.llNoPrompt = .F.
IF TYPE( "m.tlNoPrompt" ) = "L"
   m.llNoPrompt = m.tlNoPrompt
ENDIF

m.lnMaxErrCnt = 10
IF TYPE( "m.tnMaxErrCnt" ) = "N"
   DO CASE
      CASE m.tnMaxErrCnt = -1
         m.lnMaxErrCnt = m.tnMaxErrCnt
      CASE m.tnMaxErrCnt > 0
         m.lnMaxErrCnt = m.tnMaxErrCnt
      OTHERWISE
         * Nothing to do - use default value
   ENDCASE
ENDIF

*
*  Save the environment
*  --------------------
IF SET( "TALK" ) = "ON"
   SET TALK OFF
   m.lcOldTalk = "ON"
ELSE
   m.lcOldTalk = "OFF"
ENDIF

m.lcOldEsc = SET( "ESCAPE" )
SET ESCAPE OFF

m.lcOldDel  = SET( "DELETED" )
SET DELETED OFF

m.lcOldNear  = SET( "NEAR" )
SET NEAR OFF

m.lcOldError = ON( "ERROR" )
ON ERROR

*
*  If running interactively, CLEAR the screen
*  ------------------------------------------
IF m.llNoPrompt = .F.
   CLEAR
ENDIF

*
*  Do some initialization
*  ----------------------
m.lnSecStart = SECONDS()    && Time the test starts
m.gnRecordCnt  = 0
m.gnDeletedCnt = 0
m.gnSeekErrCnt = 0
m.gnForErrCnt  = 0
m.gnDupKeyCnt  = 0
m.gnScanErrCnt = 0
m.gnFoxErrCnt  = 0
m.lcAlias = ""
m.llOk = .T.

*
*  Check for a table path
*  ----------------------
IF m.llOK = .T.
   IF TYPE( "m.tcTablePath" ) != "C"
      ? "No table path was specified."
      m.llOk = .F.
   ELSE
      *< RAS 01-Feb-2015, moved to main loop in ToolCode()
      *< tcTablePath = ALLTRIM(m.tcTablePath)
      *< ? "Checking indexes for table: " + m.tcTablePath
      *< ?
   ENDIF
ENDIF

*
*  Open the specified table
*  ------------------------
IF m.llOK = .T.
   m.lnErrorNo = 0
   
   TRY 
      * RAS 04-Feb-2015, Just in case developer has the file opened already in the IDE with the same alias.
      USE IN (SELECT(JUSTSTEM(m.tcTablePath)))
      
      SELECT 0
      USE (m.tcTablePath) NOUPDATE
   
   CATCH TO loException
      m.lnErrorNo = loException.ErrorNo
      
   ENDTRY
   
   IF m.lnErrorNo != 0
      ? "ERROR: Unable to open the table '" + m.tcTablePath + "'."
      ? "Error" + STR( m.lnErrorNo )
      ? MESSAGE()
      ?
      m.llOK = .F.
   ELSE
      m.lcAlias = ALIAS()
      m.lnRecCnt = RECCOUNT( m.lcAlias )
      m.gnRecordCnt = m.lnRecCnt

      * RAS 01-Feb-2015, Added some details/attributes about the table 
      ? "Source Name:", TRANSFORM(CURSORGETPROP("SourceName", m.lcAlias))

      lcCGPDatabase = CURSORGETPROP("Database", m.lcAlias)
      
      IF EMPTY(m.lcCGPDatabase)
         ? "Is a Free Table" 
      ELSE
         ? "Database:", m.lcCGPDatabase
      ENDIF 
      
      ?
      ? "Source Type:   ", TRANSFORM(CURSORGETPROP("SourceType", m.lcAlias))
      ? "Table Type:    ", TableType(SYS(2029, m.lcAlias))
      ? "Last Updated:  ", LUPDATE(m.lcAlias)

      * Show file sizes
      lcOldCompatible = SET("Compatible")
      SET COMPATIBLE ON 

      lnDBFSize = FSIZE(FORCEEXT(m.lcAlias, "dbf"))
      ? "DBF File Size: ", ALLTRIM(TRANSFORM(lnDBFSize, "999,999,999,999")) + SPACE(1) + FileSizeWarningCheck(lnDBFSize)
      
      TRY 
         lnCDXSize = FSIZE(FORCEEXT(m.lcAlias, "cdx"))
         ? "CDX File Size: ", ALLTRIM(TRANSFORM(lnCDXSize, "999,999,999,999")) + SPACE(1) + FileSizeWarningCheck(lnCDXSize)

      CATCH TO loException
         * Ignore, there is nothing to do if there is no file available
         
      ENDTRY
      
      TRY
         lnFPTSize = FSIZE(FORCEEXT(m.lcAlias, "fpt")) 
         ? "FPT File Size: ", ALLTRIM(TRANSFORM(lnFPTSize , "999,999,999,999")) + SPACE(1) + FileSizeWarningCheck(lnFPTSize)
      
      CATCH TO loException      
         * Ignore, there is nothing to do if there is no file available
         
      ENDTRY

      SET COMPATIBLE &lcOldCompatible
      
      ? "Header Size: ", TRANSFORM(HEADER(m.lcAlias))
      ? "Record Size: ", TRANSFORM(RECSIZE(m.lcAlias))
      ? "Field Count: ", TRANSFORM(FCOUNT(m.lcAlias))
      ?

      m.lnTagCnt = TAGCOUNT( "", m.lcAlias )
      IF m.lnTagCnt = 0
         ? "No indexes are defined for the table '" + m.tcTablePath + "'."
         ?
         m.llOk = .F.
      ENDIF
   ENDIF
ENDIF


*
*  Warn the user if any of the index expressions use the FoxPro function DTOC().
*  Working with such indexes can be unreliable since the result of the index
*  expression is dependent upon current values of SET CENTURY and SET DATE.
*  -----------------------------------------------------------------------------
IF m.llOK = .T.
   * RAS 01-Feb-2015, Added tag details.
   lcTagDetails = TableTagDetails(m.lcAlias)
   
   IF EMPTY(lcTagDetails)
      * Nothing to do
   ELSE
      ? lcTagDetails
   ENDIF 

   FOR m.i = 1 TO m.lnTagCnt

      * Get the information for the next index
      m.lcTagName   = ALLTRIM(TAG( m.i, m.lcAlias ))
      m.lcIndexExpr = KEY( m.i, m.lcAlias )
      m.lcForExpr   = FOR( m.i, m.lcAlias )
      
      llWarningRecorded = .F.

      IF "DTOC(" $ m.lcIndexExpr + m.lcForExpr
         ? "Tag:", m.lcTagName, " with Expr: [" + ALLTRIM(m.lcIndexExpr) + "] ", "WARNING: This index uses the function DTOC(), problem depending on SET CENTURY"
         llWarningRecorded = .T.
         *< RAS 31-Jan-2015, removed as it is logged, no need to stop process
         *< READ
      ENDIF

      * RAS 31-Jan-2015, added additional check for variable sized indexes
      IF "ALLTRIM(" $ m.lcIndexExpr + m.lcForExpr OR "ALLT(" $ m.lcIndexExpr + m.lcForExpr
         ? "Tag:", m.lcTagName, " with Expr: [" + ALLTRIM(m.lcIndexExpr) + "] ", "WARNING: This index uses the function ALLTRIM(), make sure padded to fix length"
         llWarningRecorded = .T.
      ENDIF

      * RAS 31-Jan-2015, added additional check for variable sized indexes
      IF "RTRIM(" $ m.lcIndexExpr + m.lcForExpr OR "RTRI(" $ m.lcIndexExpr + m.lcForExpr
         ? "Tag:", m.lcTagName, " with Expr: [" + ALLTRIM(m.lcIndexExpr) + "] ", "WARNING: This index uses the function RTRIM(), make sure padded to fix length"
         llWarningRecorded = .T.
      ENDIF

      * RAS 31-Jan-2015, added additional check for variable sized indexes
      IF "STR(" $ m.lcIndexExpr + m.lcForExpr 
         ? "Tag:", m.lcTagName, " with Expr: [" + ALLTRIM(m.lcIndexExpr) + "] ", "WARNING: This index uses the function STR(), make sure it has second parameter to fix length"
         llWarningRecorded = .T.
      ENDIF

      * RAS 4-Feb-2015, added additional check for FPD/FPW "UNIQUE" Tags
      IF UNIQUE("", m.i) = .T.
         ? "Tag:", m.lcTagName, " with Expr: [" + ALLTRIM(m.lcIndexExpr) + "] ", "WARNING: This index uses obsolete UNIQUE clause, this does not guarantee uniqueness of column"
         llWarningRecorded = .T.
      ENDIF

   ENDFOR
   
   * RAS 01-Feb-2015, added white space only if warning recorded
   IF llWarningRecorded
      ?
   ENDIF 
ENDIF
   
*
*  Determine the number of deleted records
*  ---------------------------------------
IF m.llOK = .T.
   SELECT (m.lcAlias)
   ? "Determining the number of deleted records... "
   CALCULATE CNT() FOR DELETED() TO m.gnDeletedCnt
   ?? Num2Str( m.gnDeletedCnt )
ENDIF

*
*  Pass 1. Do the following:
*     - Use each defined index in turn.
*     - For each selected index, position to each record in the table.
*     - For each record in the table determine the SEEK value corresponding to that record.
*     - Using the SEEK value, perform a SEEK on the table using the current index.
*     - Check whether the SEEK was successful. 
*     - If the SEEK was successful, check for a duplicate key.
*     - Insert the result into the Pass 1 result cursor.
*  -----------------------------------------------------------------------------------------
IF m.llOK = .T.

   * Display the table path, pass number, number of indexes, and number of records
   ?
   ? "PASS 1: SEEK test"
   ? "Checking " + Num2Str( m.lnTagCnt ) + " tags on " + ;
      Num2Str( m.lnRecCnt ) + " records."

   * Scan through all defined indexes for the table
   m.llCanceled = .F.
   m.llStopTest = .F.
   FOR m.i = 1 TO m.lnTagCnt

      m.lnIdxErrCnt = 0
      m.llFoxError = .F.

      * Stop tag loop if a "cancel" request occurred
      IF m.llCanceled = .T.
         EXIT
      ENDIF

      * Select the next index
      m.lcTagName = TAG( m.i, m.lcAlias )
      m.lcIndexExpr = KEY( m.i, m.lcAlias )
      m.lcForExpr = FOR( m.i, m.lcAlias )
      
      *< RAS 04-Feb-2015, not passing parameters in VFP always returns .T. if any tag is unique, primary or candidate
      *< From FPD Topic in Help file:
      *<     If the optional parameters are omitted, UNIQUE() returns the UNIQUE status for the master controlling
      *<     index tag or index file. UNIQUE() returns false (.F.) if a master controlling index tag or index
      *<     file isn't in effect (the table is in natural record number order).
      *< m.llUnique = UNIQUE()
      m.llUnique = UNIQUE("", m.i)
      
      * RAS 04-Feb-2015, added logic to determine uniqueness violations on primary and candidate tags
      m.llPrimary = PRIMARY(m.i, m.lcAlias)
      m.llCandidate = CANDIDATE(m.i, m.lcAlias)
      
      *< RAS 04-Feb-2015, added other types if indexes to the list
      *< m.lcIndexType = IIF( m.llUnique = .T., "Unique", "Regular" )
      *< ? "Tag:", m.lcTagName, IIF( m.llUnique = .T., m.lcIndexType, "" )
      m.lcIndexType = ICASE( m.llPrimary, "Primary", m.llCandidate, "Candidate", m.llUnique, "Unique", "Regular" )
      ? "Tag:", PADR(m.lcTagName, 15), "(" + m.lcIndexType + ")"
      
      SELECT (m.lcAlias)
      SET ORDER TO (m.lcTagName)

      * Scan through all records in the table
      FOR m.lnRecNo = 1 TO m.lnRecCnt

         * Process user "cancel" request
         IF INKEY() = 27    && Stop if ESC
            m.llCanceled = .T.
         ENDIF

         * Stop the record loop, if needed
         IF m.llCanceled = .T. OR m.llStopTest = .T.
            EXIT
         ENDIF

         * Stop record loop if too many errors were detected
         IF m.lnMaxErrCnt > 0 AND m.lnIdxErrCnt >= m.lnMaxErrCnt
            EXIT
         ENDIF

         * Give the user some feedback
         IF m.lnRecNo = 1 OR m.lnRecNo % 10000 = 0
            WAIT WINDOW "Table: " + tcTablePath + SPACE(4) + "Tag: " + m.lcTagName + SPACE(4) + ;
               Num2Str( m.lnRecNo ) + "  of  " + Num2Str( m.lnRecCnt ) NOWAIT
         ENDIF

         * Initialize some variables
         m.llError = .F.
         m.lcMessage = ""
         m.lcDupKey = ""
         m.lcForErr = ""

         * Position to the next record and get the seek key and index filter value.
         * If there is no index filter expression, set the index filter value to .T.
         * to indicate that all records are included in the index.
         GO RECORD m.lnRecNo
         m.luSeekKey = EVALUATE( m.lcIndexExpr )
         m.llForValue = IIF( EMPTY( m.lcForExpr ), .T., EVALUATE( m.lcForExpr ) )

         * Position to EOF to ensure we aren't on a valid record
         GO BOTTOM
         SKIP

         * Perform the SEEK
         *
         * Set an ON ERROR to trap at least the following error:
         *    - 2066  Index file "file" is corrupted. Please rebuild it.
         *      Note: The above error became available in VFP 9. In FPD 2.6, a
         *      corrupted index can cause FoxPro to hang (that was the observed
         *      behavior in one test).
         m.lnErrorNo = 0
         ON ERROR m.lnErrorNo = ERROR()
         m.llFound = SEEK( m.luSeekKey )
         ON ERROR

         * Check for a SEEK error
         DO CASE

            * Error: A FoxPro error occurred. Stop the test.
            CASE m.lnErrorNo != 0
               m.lcMessage = "Fox Error: " + TRANSFORM( m.lnErrorNo )
               m.llError = .T.
               m.gnFoxErrCnt = m.gnFoxErrCnt + 1
               m.llError    = .T.
               m.llFoxError = .T.
               m.llStopTest = .T.

            * Okay: Didn't find the record because the FOR condition suppressed it
            CASE m.llFound = .F. AND m.llForValue = .F.
               * nothing to do

            * Error: Didn't find the record, but should have
            CASE m.llFound = .F. AND m.llForValue = .T.
               m.lcMessage = "SEEK Failed"
               m.llError = .T.

            * Error: Found the record, but the FOR condition should have suppressed it
            CASE m.llFound = .T. AND m.llForValue = .F.
               m.lcMessage = "SEEK Error, FOR"
               m.llError = .T.
               m.lcForErr = "X"
               m.gnForErrCnt = m.gnForErrCnt + 1

            * Error: Found the record, but record values did not match SEEK expression
            CASE m.llFound = .T. AND !( EVALUATE( m.lcIndexExpr )== m.luSeekKey )
               m.lcMessage = "Idx value error"
               m.llError = .T.
            
            * RAS 04-Feb-2015, added check for dupicate on Primary tag 
            * Error: Found the record, but detected a duplicate key for a unique index
            CASE m.llFound = .T. AND RECNO( m.lcAlias ) != m.lnRecNo AND m.llPrimary = .T.
               m.lcMessage = "Dup, Primary key"
               m.llError = .T.

            * RAS 04-Feb-2015, added check for dupicate on Candidate tag
            * Error: Found the record, but detected a duplicate key for a unique index
            CASE m.llFound = .T. AND RECNO( m.lcAlias ) != m.lnRecNo AND m.llCandidate = .T.
               m.lcMessage = "Dup, Candidate key"
               m.llError = .T.

            * Error: Found the record, but detected a duplicate key for a unique index
            CASE m.llFound = .T. AND RECNO( m.lcAlias ) != m.lnRecNo AND m.llUnique = .T.
               m.lcMessage = "Dup, Unique key"
               m.llError = .T.

            * Okay: Found the record, but detected a duplicate key for a non-unique index
            CASE m.llFound = .T. AND RECNO( m.lcAlias ) != m.lnRecNo AND m.llUnique = .F.
               m.lcDupKey = "X"
               m.gnDupKeyCnt = m.gnDupKeyCnt + 1

            * Okay: Found the record and nothing unusual detected
            OTHERWISE
               * nothing to do

         ENDCASE

         * If an error was detected, bump the error count
         IF m.llError = .T.
            ? " - ERROR: ", m.lnRecNo, m.lcTagName, m.luSeekKey, m.lcMessage
            m.gnSeekErrCnt = m.gnSeekErrCnt + 1
            m.lnIdxErrCnt = m.lnIdxErrCnt + 1
         ENDIF

         * If an error was detected, add it to the result cursor
         IF m.llError = .T.
            m.lcSeekKey = TRANSFORM( m.luSeekKey, "" )
            INSERT INTO curPass1 VALUES ;
                ( m.tcTablePath, ;
                  m.lnRecNo, ;
                  m.lcTagName, ;
                  m.lcIndexType, ;
                  m.lcSeekKey, ;
                  m.lcMessage, ;
                  m.lcDupKey, ;
                  m.lcForErr )
         ENDIF

      ENDFOR

      * Add final information for this tag into the Pass 1 result cursor
      DO CASE

         CASE m.llCanceled = .T.
            m.lcMessage = "User canceled"

         CASE m.llStopTest = .T.
            m.lcMessage = "Fatal error"

         CASE m.lnMaxErrCnt > 0 AND m.lnIdxErrCnt >= m.lnMaxErrCnt
            m.lcMessage = "Reached max errors"

         OTHERWISE
            m.lcMessage = "Done"

      ENDCASE

      INSERT INTO curPass1 VALUES ;
          ( m.tcTablePath, ;
            m.lnRecNo, ;
            m.lcTagName, ;
            m.lcIndexType, ;
            "", ;
            m.lcMessage, ;
            "", ;
            "" )

   ENDFOR
   WAIT CLEAR

   IF m.llCanceled = .T.
      ?
      ? "User canceled SEEK test."
      ?
   ENDIF
ENDIF

*
*  Pass 2. Do the following:
*     - Create a temporary cursor called "curRecChk" with as many blank records as
*       there are records in the specified table.
*     - Use each defined index in turn.
*     - For each selected index, add a record to the Pass 2 result cursor and then
*       SCAN the entire table.
*     - For each record selected in the SCAN, mark the corresponding record in the
*       temporary cursor to indicate the record was found.
*     - If any record is detected more than once, terminate the SCAN.
*     - If the SCAN was detected the same record twice, or if the scan did not pass
*       through every record in the table, update the Pass 2 result cursor to
*       indicate that fact.
*  --------------------------------------------------------------------------------
IF m.llOK = .T.

   * Create the record check cursor "curRecChk"
   CREATE CURSOR curRecChk ;
      ( cRecFound  C(1) )
   FOR m.i = 1 TO m.lnRecCnt
      APPEND BLANK
   ENDFOR

   * Display the table path, pass number, number of indexes, and number of records
   ?
   ? "PASS 2: SCAN test"
   ? "Checking " + Num2Str( m.lnTagCnt ) + " tags on " + ;
      Num2Str( m.lnRecCnt ) + " records."

   * Scan through all defined indexes for the table
   m.llCanceled = .F.
   m.llStopTest = .F.
   FOR m.i = 1 TO m.lnTagCnt

      * Set up to trap any FoxPro errors during the test
      m.lnErrorNo = 0
      ON ERROR m.lnErrorNo = ERROR()

      * Stop tag loop if a "cancel" request occurred
      IF m.llCanceled = .T.
         EXIT
      ENDIF

      * If this is not the first index, clear all marks in the record check cursor
      IF m.i > 1
         SELECT curRecChk
         REPLACE ALL cRecFound WITH " "
      ENDIF

      * Select the next index
      SELECT (m.lcAlias)
      m.lcTagName = TAG( m.i, m.lcAlias )
      m.lcIndexExpr = KEY( m.i, m.lcAlias )
      m.lcForExpr = FOR( m.i, m.lcAlias )
      m.llDescend = DESCENDING( "", m.i )
      
      *< RAS 04-Feb-2015, not passing parameters in VFP always returns .T. if any tag is unique, primary or candidate
      *< From FPD Topic in Help file:
      *<     If the optional parameters are omitted, UNIQUE() returns the UNIQUE status for the master controlling
      *<     index tag or index file. UNIQUE() returns false (.F.) if a master controlling index tag or index
      *<     file isn't in effect (the table is in natural record number order).
      *< m.llUnique = UNIQUE()
      m.llUnique = UNIQUE("", m.i)

      * RAS 04-Feb-2015, added logic to determine uniqueness violations on primary and candidate tags
      m.llPrimary = PRIMARY(m.i, m.lcAlias)
      m.llCandidate = CANDIDATE(m.i, m.lcAlias)
      
      *< RAS 04-Feb-2015, added other types if indexes to the list
      *< m.lcIndexType = IIF( m.llUnique = .T., "Unique", "Regular" )
      *< ? "Tag:", m.lcTagName, IIF( m.llUnique = .T., m.lcIndexType, "" )
      m.lcIndexType = ICASE( m.llPrimary, "Primary", m.llCandidate, "Candidate", m.llUnique, "Unique", "Regular" )

      ? "Tag:", PADR(m.lcTagName, 15), "(" + m.lcIndexType + ")"
      SET ORDER TO (m.lcTagName)
      GO TOP

      * Initialize some variables
      m.lnLoopCnt = 0
      m.lnP2RecCnt = 0
      m.lcMessage = ""
      m.lnP2LastRec = 0
      m.luP2LastKey = EVALUATE( m.lcIndexExpr )

      * Scan through all records in the table
      m.llFoxError = .F.
      SCAN

         m.llError    = .F.

         * Count the number of times through the loop
         m.lnLoopCnt = m.lnLoopCnt + 1

         * Process user "cancel" request
         IF INKEY() = 27    && Stop if ESC
            m.llCanceled = .T.
         ENDIF

         * Stop the record loop, if needed
         IF m.llCanceled = .T. OR m.llStopTest = .T.
            EXIT
         ENDIF

         * Give the user some feedback
         IF m.lnLoopCnt = 1 OR m.lnLoopCnt % 10000 = 0
            WAIT WINDOW "Table: " + tcTablePath + SPACE(4) + "Tag: " + m.lcTagName + SPACE(4) + ;
               Num2Str( m.lnLoopCnt ) + "  of  " + Num2Str( m.lnRecCnt ) NOWAIT
         ENDIF

         * Get some information about the current record
         m.lnRecNo = RECNO()
         m.luSeekKey = EVALUATE( m.lcIndexExpr )

         * Check this record for errors. If an error is found, flag the error
         * and terminate the SCAN, otherwise mark the record as having been found.
         DO CASE

            * Error: A FoxPro error occurred. Stop the test.
            CASE m.lnErrorNo != 0
               m.lcMessage = "Fox Error: " + TRANSFORM( m.lnErrorNo )
               m.gnFoxErrCnt = m.gnFoxErrCnt + 1
               m.llError    = .T.
               m.llFoxError = .T.
               m.llStopTest = .T.

            * Check for beginning of file
            CASE BOF()
               m.lcMessage = "BOF during SCAN"
               m.llError = .T.

            * Check for end of file
            CASE EOF()
               m.lcMessage = "EOF during SCAN"
               m.llError = .T.

            * Check for an invalid record number
            CASE !BETWEEN( m.lnRecNo, 1, m.lnRecCnt )
               m.lcMessage = "Invalid record number: " + Num2Str( m.lnRecNo )
               m.llError = .T.

            * Check for this record's index value being in sequence
            CASE m.llDescend = .F. AND m.luSeekKey < m.luP2LastKey
               m.lcMessage = "Index value < prior. Prior rec: " + Num2Str( m.lnP2LastRec )
               m.llError = .T.

            CASE m.llDescend = .T. AND m.luSeekKey > m.luP2LastKey
               m.lcMessage = "Index value > prior. Prior rec: " + Num2Str( m.lnP2LastRec )
               m.llError = .T.

            * Check for the same record being encountered twice during the SCAN.
            * If the same record is found, stop the test to avoid an infinite loop.
            OTHERWISE
               GO RECORD m.lnRecNo IN curRecChk
               IF curRecChk.cRecFound != " "
                  m.lcMessage = "Record found twice, prior was: " + Num2Str( m.lnP2LastRec )
                  m.llError = .T.
                  m.llStopTest = .T.
               ELSE
                  REPLACE curRecChk.cRecFound WITH "X"
                  m.lnP2RecCnt = m.lnP2RecCnt + 1
               ENDIF 

         ENDCASE

         * If an error was detected, bump the error count
         IF m.llError = .T.
            ? " - ERROR: ", m.lnRecNo, m.lcTagName, m.luSeekKey, m.lcMessage
            m.gnScanErrCnt = m.gnScanErrCnt + 1
         ENDIF

         * If an error was detected, add it to the result cursor
         IF m.llError = .T.
            m.lcSeekKey = TRANSFORM( m.luSeekKey, "" )
            INSERT INTO curPass2 VALUES ;
               ( m.tcTablePath, ;
                 m.lnRecNo, ;
                 m.lcTagName, ;
                 m.lcIndexType, ;
                 m.lcSeekKey, ;
                 m.lcMessage )
         ENDIF

         * Save information about this record as the last (prior) record
         m.lnP2LastRec = m.lnRecNo
         m.luP2LastKey = m.luSeekKey

      ENDSCAN

      * Add final information for this tag into the Pass 2 result cursor
      DO CASE

         CASE m.llCanceled = .T.
            m.lcMessage = "User canceled"

         CASE m.llStopTest = .T.
            m.lcMessage = "Fatal error"

         CASE m.lnP2RecCnt != m.lnRecCnt
            m.lcMessage = "Found " + Num2Str( lnP2RecCnt ) + " records out of " + ;
               Num2Str( m.lnRecCnt )
            gnScanErrCnt = gnScanErrCnt + 1

         OTHERWISE
            m.lcMessage = "Done"

      ENDCASE

      INSERT INTO curPass2 VALUES ;
         ( m.tcTablePath, ;
           m.lnRecNo, ;
           m.lcTagName, ;
           m.lcIndexType, ;
           "", ;
           m.lcMessage )

   ENDFOR
   WAIT CLEAR

   IF m.llCanceled = .T.
      ?
      ? "User canceled SCAN test."
   ENDIF

ENDIF

*
*  Go to the top of the Pass 1 and Pass 2 cursors so we'll be on the first
*  record in case the user wants to browse them
*  -----------------------------------------------------------------------
*< RAS 31-Jan-2015, moved to calling method since it does not make sense to do it before complete analysis
*<   IF m.llOK = .T.
*<      GO TOP IN curPass1
*<      GO TOP IN curPass2
*<   ENDIF

*
*  Display summary statistics and (if not suppressed) ask the user whether
*  to view the result, suspend, or terminate
*  -----------------------------------------------------------------------
IF m.llOK = .T.

   * Calculate the elapsed time for the test
   m.gnElapsedTime = ( SECONDS() + 86400 - m.lnSecStart ) % 86400

   DO WHILE .T.

      * Display summary statistics
      ?
      ? "ELAPSED TIME (seconds): " + Num2Str( m.gnElapsedTime, 1 )
      ?
      ? "RESULTS: " + m.tcTablePath
      ? "Pass 1:  " + Num2Str( m.gnSeekErrCnt ) + " SEEK errors were detected during a scan of " + ;
           Num2Str( m.lnTagCnt ) + " tags on " + Num2Str( m.lnRecCnt ) + " records."
      ? "Pass 1:  " + Num2Str( m.gnDupKeyCnt ) + " Duplicate keys were detected."
      ? "Pass 2:  " + Num2Str( m.gnScanErrCnt ) + " SCAN errors were detected during a scan of " + ;
           Num2Str( m.lnTagCnt ) + " tags on " + Num2Str( m.lnRecCnt ) + " records."
      ? "Pass n:  " + Num2Str( m.gnFoxErrCnt ) + " FoxPro errors were trapped."

      * If not suppressed, prompt the user for what to do
      IF m.llNoPrompt = .T.
         ?
         EXIT
      ELSE
         ? "Press 1 or 2 to view the Pass 1/2 results, S to suspend, or ESC to terminate."
         m.lnKey = INKEY(0)

         DO CASE

            * View Pass 1 results
            CASE CHR( m.lnKey ) = "1"
               SELECT curPass1
               BROWSE NORMAL NOMODIFY NOAPPEND NODELETE

            * View Pass 2 results
            CASE CHR( m.lnKey ) = "2"
               SELECT curPass2
               BROWSE NORMAL NOMODIFY NOAPPEND NODELETE

            * Suspend
            CASE UPPER( CHR( m.lnKey ) ) = "S"
               SUSPEND

            * Terminate
            CASE m.lnKey = 27    && ESC
               ?
               EXIT

            * Invalid response, ignore
            OTHERWISE
               * Ignore

         ENDCASE

      ENDIF

   ENDDO
ELSE 
   * RAS 04-Feb-2015, added logic as failures crushed inserts into table on data type mismatch.
   m.gnElapsedTime = ( SECONDS() + 86400 - m.lnSecStart ) % 86400
ENDIF

* RAS 01-Feb-2015, insert general statistics for this one table
INSERT INTO curGeneral VALUES ;
   ( m.tcTablePath, ;
     m.gnRecordCnt, ;
     m.gnDeletedCnt, ;
     m.gnSeekErrCnt, ;
     m.gnForErrCnt, ;
     m.gnDupKeyCnt, ;
     m.gnScanErrCnt, ;
     m.gnFoxErrCnt, ;
     m.gnElapsedTime  )


*
*  Clean up
*  --------
SET TALK &lcOldTalk
SET ESCAPE &lcOldEsc
SET DELETED &lcOldDel
SET NEAR &lcOldNear
ON ERROR &lcOldError

IF USED( "curRecChk" )
   USE IN curRecChk
ENDIF

IF USED( m.lcAlias )
   USE IN (m.lcAlias)
ENDIF

*
*  SELECT the Pass 1 cursor before returning
*  -----------------------------------------
IF USED( "curPass1" )
   SELECT curPass1
ENDIF

*
*  Return to caller
*  ----------------
llReturnVal = IIF( m.gnSeekErrCnt + m.gnScanErrCnt > 0, -1, 0 )


SET ALTERNATE OFF 
SET ALTERNATE TO 
    
SET SAFETY &lcOldSafety

RETURN llReturnVal



******************************************************************************
*  Internal function to return a numeric string in the format "999,999,999".
******************************************************************************
FUNCTION Num2Str(tnValue, tnDec)

PRIVATE lcDecFmt
PRIVATE lcFuncRtn

IF TYPE( "m.tnDec" ) != "N" OR !BETWEEN( m.tnDec, 1, 10 )
   m.lcDecFmt = ""
ELSE
   m.lcDecFmt = "." + REPLICATE( "9", m.tnDec )
ENDIF

IF TYPE( "m.tnValue" ) != "N"
   m.lcFuncRtn = ""
ELSE
   m.lcFuncRtn = ALLTRIM( TRANSFORM( m.tnValue, "999,999,999,999" + m.lcDecFmt ) )
ENDIF

RETURN m.lcFuncRtn


********************************************************************************
*  METHOD NAME: TableTagDetails
*
*  AUTHOR: Richard A. Schummer, February 2015
*
*  METHOD DESCRIPTION:
*    
*    
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE TableTagDetails(tcAlias)

#DEFINE ccINDEXTAG_LITERAL          "Index Tags: "
#DEFINE ccTAGNAME_LITERAL           "Tag Name: "
#DEFINE ccTAGTYPE_LITERAL           " - Type: "
#DEFINE ccTAGKEYEXPR_LITERAL        " - Key Expression: "
#DEFINE ccTAGLENGTH_LITERAL         " - Length: "
#DEFINE ccTAGFILTER_LITERAL         " - Filter: "
#DEFINE ccTAGORDER_LITERAL          " - Order: "
#DEFINE ccTAGCOLLATE_LITERAL        " - Collate Sequence: "
#DEFINE ccTAGNOFILTER               "(nothing)"

LOCAL lcIndexList, ;
      llTagDeleted, ;
      lnTags, ;
      lnY

lnTags = ATAGINFO(laTag, SPACE(0), tcAlias)

IF lnTags > 0
   lcIndexList = ccCRLF + ccINDEXTAG_LITERAL + ccCRLF 
   
   FOR m.lnY = 1 TO m.lnTags
      lcIndexList = lcIndexList + ; 
                    ListPADL(TRANSFORM(m.lnY) + ". ", 4) + ;
                    ListPADL(ccTAGNAME_LITERAL, LENC(ccTAGNAME_LITERAL) + 1) + ;
                    ListPADL(laTag[m.lnY, 1], LENC(laTag[m.lnY, 1]) + 1) + ccCRLF
      lcIndexList = lcIndexList + ; 
                    ListPADL(ccTAGTYPE_LITERAL, LENC(ccTAGTYPE_LITERAL) + 1) + ;
                    ListPADL(LOWER(laTag[m.lnY, 2]), LENC(laTag[m.lnY, 2]) + 1) + ccCRLF
      lcIndexList = lcIndexList + ; 
                    ListPADL(ccTAGKEYEXPR_LITERAL, LENC(ccTAGKEYEXPR_LITERAL) + 1) + ;
                    ListPADL(laTag[m.lnY, 3], LENC(laTag[m.lnY, 3]) + 1) + ccCRLF

      * RAS 01-Feb-2015, added key length to allow developer to see if it is close to limits.
      lcIndexList = lcIndexList + ; 
                    ListPADL(ccTAGLENGTH_LITERAL, LENC(ccTAGLENGTH_LITERAL) + 1) + ;
                    TRANSFORM(LENC(TRANSFORM(EVALUATE(laTag[m.lnY, 3]))), "999") + ccCRLF    
                    
      * RAS 28-Nov-2009, check to see if we have a deleted tag to quickly count the deleted records
      IF "deleted()" $ LOWER(laTag[m.lnY, 3])
         llTagDeleted = .T. 
      ENDIF             
                    
      lcIndexList = lcIndexList + ; 
                    ListPADL(ccTAGFILTER_LITERAL, LENC(ccTAGFILTER_LITERAL) + 1) + ;
                    ListPADL(IIF(EMPTY(laTag[m.lnY, 4]), ccTAGNOFILTER, ;
                                                         LOWER(laTag[m.lnY, 4])), MAX(LENC(laTag[m.lnY, 4]), 9) + 1) + ccCRLF
      lcIndexList = lcIndexList + ; 
                    ListPADL(ccTAGORDER_LITERAL, LENC(ccTAGORDER_LITERAL) + 1) + ;
                    ListPADL(LOWER(laTag[m.lnY, 5]), LENC(laTag[m.lnY, 5]) + 1) + ccCRLF
      lcIndexList = lcIndexList + ; 
                    ListPADL(ccTAGCOLLATE_LITERAL, LENC(ccTAGCOLLATE_LITERAL) + 1) + ;
                    ListPADL(LOWER(laTag[m.lnY, 6]), LENC(laTag[m.lnY, 6]) + 1) + ccCRLF + ccCRLF
   ENDFOR 
ELSE
   * Nothing to do
ENDIF 
 
RETURN lcIndexList


********************************************************************************
*  METHOD NAME: FileSizeWarningCheck
*
*  AUTHOR: Richard A. Schummer, February 2015
*
*  METHOD DESCRIPTION:
*    Process the file sizes of the different data files to see if they are getting
*    close to the Visual FoxPro 2GB limits.
*
*  INPUT PARAMETERS:
*    tnFileSize = numeric, required, byte size of the file to be checked.
* 
*  OUTPUT PARAMETERS:
*    lcFileSizeStatus = character, output of the file size status.
* 
********************************************************************************
PROCEDURE FileSizeWarningCheck(tnFileSize)

LOCAL lcFileSizeStatus, ;
      lnFileLimitWarningLevel, ;
      lnPercentage

tnFilesize    = tnFilesize * 1.0000000000
lnTwoGigLimit = cnTWOGIGLIMIT

lnOldDecimals = SET("Decimals")

SET DECIMALS TO 10

IF VARTYPE(tnFileSize) = "N"
   lnFileLimitWarningLevel = ROUND(cnTWOGIGLIMIT * cnHOWCLOSETOLIMIT, 0)
   
   * The reason for dividing each value by 1024 is to get the math away from the 2GB limit for integers.
   lnPercentage            = ROUND((tnFileSize/1024) / (lnTwoGigLimit/1024), 10) * 100
   
   lcFileSizeStatus = "(" + ALLTRIM(TRANSFORM(lnPercentage, "999.99")) + "% of VFP limit"
   
   IF  lnFileLimitWarningLevel > tnFileSize
      * All is good in the world, nothing else to report
   ELSE
      * Raise a flag that file is getting close to VFP limits
      lcFileSizeStatus = lcFileSizeStatus +  ", WARNING: getting close to limit"
   ENDIF

   lcFileSizeStatus = lcFileSizeStatus + ")"
ELSE
   lcFileSizeStatus = "Valid file size not passed to checker"
ENDIF 

SET DECIMALS TO (lnOldDecimals)

RETURN lcFileSizeStatus
 

********************************************************************************
*  METHOD NAME: TableType
*
*  AUTHOR: Richard A. Schummer, February 2015
*
*  METHOD DESCRIPTION:
*    This method determines the type of table opened.
*
*  INPUT PARAMETERS:
*    tcType = Numeric, required, value determined by SYS(2029).
* 
*  OUTPUT PARAMETERS:
*    lcReturnVal = character, description of the table type.
* 
********************************************************************************
PROCEDURE TableType(tcType)

LOCAL lcReturnVal

IF PCOUNT() = 1 AND VARTYPE(tcType) = "C"
   DO CASE
      CASE tcType = "3"
         lcReturnVal = "Previous versions of FoxPro, FoxBASE+, dBASE III PLUS, and dBASE IV with no memo field."

      CASE tcType = "48"
         lcReturnVal = "Visual FoxPro with or without a memo field."

      CASE tcType = "49"
         lcReturnVal = "Visual FoxPro with an Autoinc field (Visual FoxPro 8 and greater)."

      CASE tcType = "50"
         lcReturnVal = "Visual FoxPro with a Varchar, Varbinary or Blob field (Visual FoxPro 9 and greater)."

      CASE tcType = "67"
         lcReturnVal = "dBASE IV SQL table with no memo field."

      CASE tcType = "99"
         lcReturnVal = "dBASE IV SQL System table with no memo field."

      CASE tcType = "131"
         lcReturnVal = "FoxBASE+ and dBASE III PLUS table with a memo field."

      CASE tcType = "139"
         lcReturnVal = "dBASE IV table with a memo field."

      CASE tcType = "203"
         lcReturnVal = "dBASE IV SQL table with a memo field."

      CASE tcType = "245"
         lcReturnVal = "Previous versions of FoxPro with a memo field."

      CASE tcType = "0"
         lcReturnVal = "No table open."

      OTHERWISE
         lcReturnVal = "Unknown table type."
      
   ENDCASE
ELSE
   lcReturnVal = "Invalid parameter to Table Type process."
ENDIF 

RETURN m.lcReturnVal


********************************************************************************
*  METHOD NAME: ListPADL
*
*  AUTHOR: Richard A. Schummer, April 2010
*
*  METHOD DESCRIPTION:
*    This method handles the normal PADL function, but optionally will insert 
*    a TAB character before text instead for use in tab delimted formatting 
*    like a word processing table when documenting tables in specifications.
*    
*  INPUT PARAMETERS:
*    tcString    = character, required, string to be padded on left side.
*    tnPadding   = integer, required, number of spaces to pad when not tabbing.
*    tcCharacter = character, optional, character used in padding.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ListPADL(tcString, tnPadding, tcCharacter)

LOCAL lcString

lcString = SPACE(0)

IF VARTYPE(tcCharacter) # "C"
   tcCharacter = SPACE(1)
ENDIF 

IF clTABDELIMITED
   lcString = ccTAB + ALLTRIM(tcString)
ELSE
   lcString = PADL(tcString, tnPadding, tcCharacter)
ENDIF 

RETURN lcString

*: EOF :*  