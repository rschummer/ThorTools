******************************************************************************************
*  PROGRAM: Thor_tool_CheckMemoFileIntegrity.prg
*
*  AUTHOR: Richard A. Schummer, March 2015
*
*  COPYRIGHT © 2015   All Rights Reserved.
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
*     This Thor tool is designed to check the health of the memo file for tables. There
*     are three different ways this program runs, in this priority of determining how
*     the tables for analysis are figured out.
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
*     Logs potential issues with memo fields in the tables.
*
*     Base code to check integrity of the memo fields was leveraged from the 
*
*  CALLING SYNTAX:
*     This program works fine outside of Thor, but was designed with Thor integration
*
*     DO Thor_Tool_CheckMemoFileIntegrity.prg
*     Thor_Tool_CheckMemoFileIntegrity()
*     Call from Thor menu
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
*                             C H A N G E    L O G                              
*
*    Date     Developer               Version  Description
* ----------  ----------------------  -------  ---------------------------------
* 03/19/2015  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     1.1      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

#DEFINE ccCRLF            CHR(13)+CHR(10)
#DEFINE ccTAB             CHR(9)
#DEFINE ccMsgCaption      "Memo File Integrity Checker"
#DEFINE ccLOGFILE         "FPTIntegrityCheckLog.txt"
#DEFINE ccALTERNATEFILE   "FPTIntegritySingleFile.txt"
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
      .Prompt          = 'Check Memo File Integrity'           && used in menus
      .StatusBarText   = 'Evaluate the memo files in DBCX Metadata, database container, or folder.'  
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2     && a description for the tool
         This program evaluates the memo files for tables found in DBCX Metadata, database container, or a folder.
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source        = "WLC"                      && where did this tool come from?  Your own initials, for instance
      .Category      = "WLC"                      && creates categorization of tools; defaults to .Source if empty
      .Sort          = 0                          && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version       = "Version 1.1, September 20, 2015"           && e.g., 'Version 7, May 18, 2011'
      .Author        = "Rick Schummer"
      .Link          = "https://github.com/rschummer/ThorTools"    && link to a page for this tool
      .VideoLink     = ''                                                                    && link to a video for this tool
      
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
*    check if there is memo issues in the different tables.
*
*  INPUT PARAMETERS:
*    lxParam1 = unknown type, optional, standard parameter passed in by Thor.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode
LPARAMETERS lxParam1

LOCAL lcCode, ;
      lcDBCFile, ;
      lcDBCName, ;
      lcLogText, ;
      lcOldConsole, ;
      lcOldSafety, ;
      lcSourceSpecifics, ;
      lcTableName, ;
      lcTableSource, ;
      lcTableToCheck, ;
      llForceEachTable, ;
      llMemoFile, ;
      lnDBCFiles, ;
      lnFilesInFolder, ;
      lnI, ;
      lnJ, ;
      lnOldMemoWidth, ;
      lnTables, ;
      loException AS Exception
      
llMemoFile       = .F.
llForceEachTable = .T.   && Determine if you want statistics on tables without memo file. True for all, false for memo only.

TRY
   USE IN (SELECT("CoreMeta"))

   CD ?

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
         
         lcTableSource     = ccFOLDERSOURCE
         lcSourceSpecifics = "Folder selected: " + LOWER(lcFolderForFiles)

   ENDCASE

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
               "Path: " + LOWER(SET("Path")) + ccCRLF + ccCRLF + ;
               lcSourceSpecifics + ccCRLF + ccCRLF

   lcLogText = lcLogText + ;
               "**** Search for 'ERROR:' for any issues ****" + ccCRLF + ccCRLF + ;
               ccCRLF + ccCRLF ;
   
   ************************************************************************************
   * Loop for all the tables.
   FOR lnJ = 1 TO lnTables
      * Determine the table name based on the type of source file list
      DO CASE
         * =================================
         CASE lcTableSource = ccDBCXSOURCE
            IF "!" $ laTables[lnJ]
               laTables[lnJ] = ALLTRIM(laTables[lnJ])
               
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
                       "Checking Memo File for table: " + UPPER(m.lcTableToCheck) + ccCRLF 
      
      
      * Run the inspection of the memo file for a table
      TRY 
         lnOldMemoWidth = SET("Memowidth")
         lcOldConsole   = SET("Console")
         
         SET MEMOWIDTH TO 225
         SET CONSOLE OFF 
         
         llMemoFile = InspectMemo(ALLTRIM(m.lcTableToCheck))
     
         SET CONSOLE &lcOldConsole
         SET MEMOWIDTH TO (lnOldMemoWidth)
         
      CATCH TO loException      
         lcCode = "ERROR: " + m.loException.Message + ;
                  " [" + TRANSFORM(m.loException.Details) + "] " + ;
                  " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
                  " in " + m.loException.Procedure + ;
                  " on " + TRANSFORM(m.loException.LineNo)
         
         lcLogText = lcLogText + ccCRLF + lcCode + ccCRLF 
         
         * In case the log file in InspectMemo was not closed via error
         SET ALTERNATE OFF 
         SET ALTERNATE TO 
         
         * Error might leave the table open so attempt to close it.
         USE IN (SELECT(m.lcTableToCheck))
         
      ENDTRY

      * Add in the text from the index alternate file log.
      TRY
         IF llMemoFile OR llForceEachTable
            lcLogText = lcLogText + ccCRLF + ;
                        FILETOSTR(ccALTERNATEFILE) + ;
                        ccCRLF 
         ELSE 
            * Nothing to do
         ENDIF 
      
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
*  METHOD NAME: InspectMemo
*
*  AUTHOR: Richard A. Schummer, March 2015
*
*  METHOD DESCRIPTION:
*     This method opens the specified table and tests its contents for memo 
*     corruption by opening the table and scanning the records.
*
*  INPUT PARAMETERS:
*    tcTablePath = character, required, the name of the table and path to open it.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE InspectMemo(tcTablePath)

LOCAL llReturnVal, ;
      llErrorFound, ;
      lcAlias, ;
      lcCGPDatabase, ;
      lcOldCompatible, ;
      lcOldDel, ;
      lcOldError, ;
      lcOldEsc, ;
      lcOldNear, ;
      lcOldSafety, ;
      lcOldTalk, ;
      llNoPrompt, ;
      llOk, ;
      lnCDXSize, ;
      lnDBFSize, ;
      lnErrorNo, ;
      lnFPTSize, ;
      lnRecCnt, ;
      lnSecStart, ;
      lnTagCnt, ;
      lnElapsedTime, ;
      loData, ;
      loException AS Exception

lcOldSafety = SET("Safety")
SET SAFETY OFF

SET ALTERNATE TO ccALTERNATEFILE
SET ALTERNATE ON 


*  Save the environment
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


*  Do some initialization
m.lnSecStart   = SECONDS()    && Time the test starts
m.lcAlias      = SPACE(0)
m.llOk         = .T.
m.llErrorFound = .F.


*  Check for a table path
IF m.llOK = .T.
   IF TYPE( "m.tcTablePath" ) != "C"
      ? "No table path was specified."
      m.llOk         = .F.
      m.llErrorFound = .T.
   ELSE
      * Nothing to do
   ENDIF
ENDIF

*  Open the specified table
IF m.llOK = .T.
   m.lnErrorNo = 0
   
   TRY 
      * RAS 04-Feb-2015, Just in case developer has the file opened already in the IDE with the same alias.
      USE IN (SELECT(JUSTSTEM(m.tcTablePath)))
      
      SELECT 0
      USE (m.tcTablePath) NOUPDATE
   
   CATCH TO loException
      m.lnErrorNo    = loException.ErrorNo
      m.llErrorFound = .T.
      
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

      * RAS 01-Feb-2015, Added some details/attributes about the table 
      ? "Source Name:", TRANSFORM(CURSORGETPROP("SourceName", m.lcAlias))

      lcCGPDatabase = CURSORGETPROP("Database", m.lcAlias)
      
      IF EMPTY(m.lcCGPDatabase)
         ? "Is a Free Table" 
      ELSE
         ? "Database:", m.lcCGPDatabase
      ENDIF 
      
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
      
      m.llOk = .F.      
      
      TRY
         lnFPTSize = FSIZE(FORCEEXT(m.lcAlias, "fpt")) 
         ? "FPT File Size: ", ALLTRIM(TRANSFORM(lnFPTSize , "999,999,999,999")) + SPACE(1) + FileSizeWarningCheck(lnFPTSize)
         m.llOk = .T.
      
      CATCH TO loException      
         * Ignore, there is nothing to do if there is no file available
                  
      ENDTRY

      SET COMPATIBLE &lcOldCompatible
      
      m.lnTagCnt = TAGCOUNT(SPACE(0), m.lcAlias )
      
      ? "Header Size:   ", TRANSFORM(HEADER(m.lcAlias))
      ? "Record Size:   ", TRANSFORM(RECSIZE(m.lcAlias))
      ? "Field Count:   ", TRANSFORM(FCOUNT(m.lcAlias))
      ? "Tag Count:     ", TRANSFORM(m.lnTagCnt)
      ? "Record Count:  ", ALLTRIM(TRANSFORM(RECCOUNT(m.lcAlias), "9,999,999,999"))
      ?
   ENDIF
ENDIF

IF m.llOK = .T.
   TRY
      SCAN
         SCATTER MEMO NAME loData
      ENDSCAN
      
      ? "Memo file is clean!"

   CATCH TO loException
      ? "**Error: " + PADR(TRANSFORM(loError.ErrorNo),10) + PADR(loError.Message, 50)
      m.llErrorFound = .T.

   ENDTRY
ENDIF
   
m.lnElapsedTime = ( SECONDS() + 86400 - m.lnSecStart ) % 86400

?
? "Time for memo check: " + TRANSFORM(lnElapsedTime) + " seconds"
?

*  Clean up
SET TALK &lcOldTalk
SET ESCAPE &lcOldEsc
SET DELETED &lcOldDel
SET NEAR &lcOldNear
ON ERROR &lcOldError

USE IN (SELECT(m.lcAlias))

SET ALTERNATE OFF 
SET ALTERNATE TO 
    
SET SAFETY &lcOldSafety

RETURN m.llOK OR m.llErrorFound


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