******************************************************************************************
*  PROGRAM: Thor_Tool_CheckDataIntegrity.prg
*
*  AUTHOR: Richard A. Schummer, November 2013
*
*  COPYRIGHT © 2013-2018   All Rights Reserved.
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
*     This program checks to see if the data in tables can be cleanly opened. If successful to open
*     the table, it provides lots of interesting details:
*         - Date/time of the file
*         - Tags count (and comparison to counts in DBCX metadata if it exists)
*         - Record counts
*         - Notice if the table is empty
*         - Details of tables closing in on 2GB limits if approaching
*
*  CALLING SYNTAX:
*     DO Thor_Tool_CheckDataIntegrity.prg
*     Thor_Tool_CheckDataIntegrity()
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
* 11/14/2013  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 03/22/2015  Richard A. Schummer     2.0      Added ability to leverage DBCX metadata
* ----------------------------------------------------------------------------------------
* 03/25/2015  Richard A. Schummer     2.1      Minor program cleanup
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     2.2      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
* 02/20/2016  Richard A. Schummer     2.2      Added index tag check, record count and 
*                                              empty table notation to the log.
* ----------------------------------------------------------------------------------------
* 08/19/2016  Richard A. Schummer     2.3      Fixed message when FREE table name cannot
*                                              be opened, previously did not display 
*                                              correct alias of the free table
* ----------------------------------------------------------------------------------------
* 10/14/2017  Richard A. Schummer     2.4      SET TABLEPROMPT OFF logic
* ----------------------------------------------------------------------------------------
* 02/06/2018  Richard A. Schummer     3.0      Index tag issue check for free tables  
*                                              (only was doing DBC tables), message to show
*                                              tag on its own line (more obvious),  
*                                              lowercased path and file names, added 
*                                              warnings that file sizes getting near 2GB 
*                                              limit. See setting percent using constant
*                                              cnHOWCLOSETOLIMIT default to 80%. Fixed issue
*                                              with tables already opening not opening again
*                                              giving false positive on potential problem.
* ----------------------------------------------------------------------------------------
* 05/01/2018  Richard A. Schummer     3.1      Added table file date/time, sorted free
*                                              tables alphabetically like contained tables.
*                                              Free table alias upper case like contained
*                                              tables, and path/file name lower cased like
*                                              contained tables. Optionally display DBCX
*                                              Caption for the table if DBCX metadata is
*                                              present and developer filled them in. See 
*                                              clSHOW_DBCX_CAPTION to toggle on or off. 
* ----------------------------------------------------------------------------------------
* 08/02/2020  Richard A. Schummer     3.3      Added table DBC file sizes for all three 
*                                              files. Added date/time stamp to the log
*                                              file name for situations where you might 
*                                              need to compare before and after file repairs.
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

#DEFINE ccTOOLNAME    "WLC DBC/DBF/FPT/SDT Integrity Checker"

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.

IF PCOUNT() = 1 AND 'O' = VARTYPE(lxParam1) AND 'thorinfo' == LOWER(lxParam1.Class)
   WITH lxParam1
   
      * Required
      .Prompt          = 'Check Data (DBC/DBF/FPT/SDT) Integrity' && used in menus
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2 && a description for the tool
         This program analyzes and inspects Database Container (DBC) integrity by VALIDATE and opening tables. It also uses DBCX metadata if available to check into free tables, and to make sure they can be opened without error.   
      ENDTEXT  
      
      .StatusBarText   = 'Analyze and inspect Database Container integrity by VALIDATE and opening tables.'  
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source        = "WLC"                      && where did this tool come from?  Your own initials, for instance
      .Category      = "WLC"                      && creates categorization of tools; defaults to .Source if empty
      .Sort          = 0                          && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version       = "Version 3.1, May 1, 2018"                  && e.g., 'Version 7, May 18, 2011'
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


********************************************************************************
*  METHOD NAME: ToolCode
*
*  AUTHOR: Richard A. Schummer, November 2013
*
*  METHOD DESCRIPTION:
*    Main tool process code.
*
*  INPUT PARAMETERS:
*    lxParam1 = unknown type, standard parameter passed in by Thor.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode

#DEFINE ccCRLF                 CHR(13)+CHR(10)
#DEFINE ccSPACEGAP             SPACE(5)
#DEFINE ccLOGFILEBASE          "DBCIntegrityCheckLog"
#DEFINE ccDBCVALIDATEFILE      "DBCValidate.txt"
#DEFINE cnTWOGIGLIMIT          (2*1024*1024*1024) -1
#DEFINE cnHOWCLOSETOLIMIT      .800000
#DEFINE clFILESIZEWARNINGS     .T.
#DEFINE ccFILESPECSTYLE        "SHORT"
#DEFINE clSHOW_DBCX_CAPTION    .F.

* Internationalization opportunities
#DEFINE ccRECORDSLITERAL       "Records"
#DEFINE ccTAGSLITERAL          "Tags"
#DEFINE ccSUCCESSMSG           "successfully opened..."
#DEFINE ccFAILEDMSG            " ************** FAILED..."
#DEFINE ccEMPTYTABLELITERAL    "EMPTY TABLE"
#DEFINE ccBADINDEXCOUNTLITERAL "** ERROR: bad tag count"

LOCAL loException as Exception,;
      lcLogText , ;
      lcOldSafety, ;
      lnOpened, ;
      lnFailed, ;
      lcCode, ;
      lcDBCFile, ;
      lcFreeTable, ;
      lnDBCFiles, ;
      lnFailedFree, ;
      lnFreeTables, ;
      lnI, ;
      lnOpenedFree, ;
      lnTables, ;
      lnTags, ;
      llFileSizeWarnings, ;
      ldFileTimeStamp , ;
      lcFileTimeStamp, ;
      lcOldCompatible

* Setting if you want the file size percentage warnings.
llFileSizeWarnings = clFILESIZEWARNINGS

lcOldSafety = SET("Safety")
SET SAFETY OFF

lcOldTablePrompt = SET("TablePrompt")
SET TABLEPROMPT OFF 

CD ?

lcLogText = ccTOOLNAME + ccCRLF + ;
            TRANSFORM(DATETIME()) + ccCRLF

TRY
   CloseData()
   
   lnDBCFiles = ADIR(laDatabases, "*.DBC")
   
   IF lnDBCFiles = 1
      lcDBCFile = laDatabases[1, 1]
   ELSE
      lcDBCFile = GETFILE("dbc")
   ENDIF 
   
   lcDBCFile = ALLTRIM(m.lcDBCFile)
 
   IF EMPTY(m.lcDBCFile)
      * Nothing to do
   ELSE
      WAIT WINDOW "Building temporary cursor of DBCX Metadata..." NOWAIT 
      CollectDBCXMetadata()

      * Open database
      WAIT WINDOW "Opening database, please wait..." NOWAIT 
      OPEN DATABASE (m.lcDBCFile) SHARED 
      SET DATABASE TO (JUSTSTEM(m.lcDBCFile))
      
      * RAS 02-Aug-2020, add file sizes of the three files to the log file.
      lcOldCompatible = SET("Compatible")
      SET COMPATIBLE ON
      
      lnDBCFileSize = FSIZE(m.lcDBCFile)
      lnDCXFileSize = FSIZE(FORCEEXT(m.lcDBCFile, "DCX"))      
      lnDCTFileSize = FSIZE(FORCEEXT(m.lcDBCFile, "DCT")) 
      
      SET COMPATIBLE &lcOldCompatible     
      
      lcLogText = m.lcLogText + ;
                  "DBC: " + LOWER(FULLPATH(DBC())) + ccCRLF + ;
                  "Current Folder: " + FULLPATH(CURDIR()) + ccCRLF + ccCRLF + ;
                  "DBC file size: " + ALLTRIM(TRANSFORM(m.lnDBCFileSize, "999,999,999,999")) + ccCRLF + ;
                  "DCX file size: " + ALLTRIM(TRANSFORM(m.lnDCXFileSize, "999,999,999,999")) + ccCRLF + ;
                  "DCT file size: " + ALLTRIM(TRANSFORM(m.lnDCTFileSize, "999,999,999,999")) + ;
                  ccCRLF + ccCRLF
      
      WAIT WINDOW "Determining number of tables to analyze" NOWAIT            
      lnTables  = ADBOBJECTS(laTables, "TABLE")
      
      lcLogText = m.lcLogText + ;
                  TRANSFORM(lnTables) + " Tables to check." + ;
                  ccCRLF 
      
      WAIT WINDOW "Validating database, please wait..." NOWAIT
      
      * Database container validation
      VALIDATE DATABASE NOCONSOLE TO FILE (ccDBCVALIDATEFILE)
      
      lcLogText = m.lcLogText + ;
                  FILETOSTR(ccDBCVALIDATEFILE) + ;
                  ccCRLF + ccCRLF
      
      DELETE FILE (ccDBCVALIDATEFILE) RECYCLE 

      * Check DBC tables
      lnOpened = 0
      lnFailed = 0
      
      IF lnTables > 0
         =ASORT(laTables)
         
         lnI = 1

         * Loop through all the tables
         DO WHILE m.lnI <= m.lnTables 
            TRY
               lcAlias       = JUSTSTEM(laTables[lnI])
               laTables[lnI] = ALLTRIM(laTables[lnI])
               
               WAIT WINDOW LOWER(laTables[lnI]) + " - " + TRANSFORM((m.lnI/m.lnTables)*100) + "%" NOWAIT 
               
               * Close this table just in case it is already open
               USE IN (SELECT(ALLTRIM(laTables[lnI])))
               
               * Attempt to open the table, using DBC!LongFileName provided by ADBOBJECTS()
               USE (m.lcDBCFile + "!" + laTables[lnI]) SHARED AGAIN IN 0 NOUPDATE 
               lnOpened  = m.lnOpened + 1 
               
               * Collect the DBCX Caption if using DBCX metadata
               IF clSHOW_DBCX_CAPTION
                  lcDBCXCap = GetDBCXMetaCaption(m.lcDBCFile, laTables[lnI])
               ELSE
                  lcDBCXCap = NULL
               ENDIF 
               
               IF ISNULL(m.lcDBCXCap)
                  * Nothing to do
               ELSE
                  lcDBCXCap = ALLTRIM(m.lcDBCXCap)
               ENDIF 
               
               * Log the table specifics...
               lcLogText = m.lcLogText + ;
                           PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnTables)), "0") + ")  " + ;
                           laTables[lnI] + ;
                           IIF(ISNULL(m.lcDBCXCap) OR EMPTY(m.lcDBCXCap), SPACE(0), " [" + m.lcDBCXCap + "]") + ;
                           SPACE(1) + ccSUCCESSMSG + ;
                           " (" + ;
                           LOWER(FULLPATH(DBF(laTables[lnI]))) + ;
                           ")" 

               * Get the date/time for the table.
               ldFileTimeStamp = FDATE(DBF(laTables[lnI]), 1)
               lcFileTimeStamp = TRANSFORM(m.ldFileTimeStamp)

               lcLogText = m.lcLogText + ;
                           " {" + m.lcFileTimeStamp + "} "

               * See if there are issues with tags missing
               * First, see what is live in the table
               lnTags    = ATAGINFO(laTags, SPACE(0), m.lcAlias)

               lcLogText = m.lcLogText + ;
                           ":  " + ;
                           TRANSFORM(RECCOUNT(m.lcAlias)) + SPACE(1) + ccRECORDSLITERAL + ", " + ;
                           TRANSFORM(m.lnTags) + SPACE(1) + ccTAGSLITERAL

               lcLogText = m.lcLogText + ;
                           IIF(RECCOUNT(m.lcAlias) = 0, ccSPACEGAP + ccEMPTYTABLELITERAL, "")
               
               * Second, see how many tags are in the Stonefield Database Toolkit / DBCX metadata
               lnCoreMetaTags = GetTypeCount(LOWER(JUSTSTEM(m.lcDBCFile)), LOWER(m.lcAlias) + ".", "I")
               
               IF ISNULL(m.lnCoreMetaTags)
                  * Nothing to report for the index tags because no DBCX metadata
               ELSE
                  IF m.lnTags # m.lnCoreMetaTags
                     lcLogText = m.lcLogText + ;
                                 ccCRLF + ccSPACEGAP + ; 
                                 "- " + ccBADINDEXCOUNTLITERAL + " - " + ;
                                 TRANSFORM(m.lnTags) + " table tags and " + ;
                                 TRANSFORM(m.lnCoreMetaTags) + " metadata tags" 
                  ENDIF 
               ENDIF 

               * Display file size details and determine if there are issues with file sizes
               DO CASE
                  * Only show if there are problems with the file size getting to percentage of limit
                  CASE UPPER(ccFILESPECSTYLE) = "SHORT"              
                     lcLogText = m.lcLogText + ;
                                 FileSpecificationsShort(m.lcAlias, .T.)

                  * Show regardless if there are problems
                  CASE UPPER(ccFILESPECSTYLE) = "LONG"
                     lcLogText = m.lcLogText + ;
                                 FileSpecificationsLong(m.lcAlias)

                  OTHERWISE
                     * No file specifications
                     
               ENDCASE


               lcLogText = m.lcLogText + ccCRLF

               * Close this table and move on to the next one
               USE IN (SELECT(laTables[lnI]))
               
               *? Potential optional future enhancements for the process:
               *?  - Scan tables
               *?  - Read memo
               
            CATCH TO loException   
               lnFailed  = m.lnFailed + 1 
               lcLogText = m.lcLogText + ;
                           ccCRLF + ;
                           PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnTables)), "0") + ")  " + ;
                           ALLTRIM(laTables[lnI]) + SPACE(1) + ccFAILEDMSG + m.loException.Message + ;
                           ccCRLF + ccCRLF

            ENDTRY

            lnI = m.lnI + 1 
         ENDDO
      ENDIF 
      
      * Check free tables, but only if we know which ones via Stonefield Database Toolkit / DBCX metadata.
      lnOpenedFree = 0
      lnFailedFree = 0

      IF FILE("coremeta.DBF")
         SELECT * ;
            FROM coremeta ;
            WHERE EMPTY(cDBCName) ;
              AND cRecType = "T" ;
            ORDER BY cObjectNam ;
            INTO CURSOR curFreetables
         
         IF USED("curFreetables")
            lnFreeTables = RECCOUNT("curFreetables")

            lcLogText    = m.lcLogText + ccCRLF + ;
                           "FREE Tables..." + ccCRLF
         ELSE
            lnFreeTables = 0
         ENDIF 
         
         lnI = 1
         
         * Loop through all the tables
         DO WHILE m.lnI <= m.lnFreeTables 
            TRY
               lcFreeTable = ALLTRIM(curFreetables.cObjectNam)
               lcFreeAlias = JUSTSTEM(m.lcFreeTable)
               
               WAIT WINDOW LOWER(m.lcFreeTable) + " - " + TRANSFORM((m.lnI/m.lnFreeTables)*100) + "%" NOWAIT 
               
               * Close this table just in case it is already open
               USE IN (SELECT(m.lcFreeTable))
               
               * Attempt to open the table
               USE (m.lcFreeTable) SHARED AGAIN IN 0 NOUPDATE 
               lnOpenedFree = m.lnOpenedFree + 1 

               * Collect the DBCX Caption if using DBCX metadata
               lcDBCXCap = GetDBCXMetaCaption(SPACE(0), m.lcFreeTable)
               
               IF ISNULL(m.lcDBCXCap)
                  * Nothing to do
               ELSE
                  lcDBCXCap = ALLTRIM(m.lcDBCXCap)
               ENDIF 
               
               * Log the table specifics...
               lcLogText    = m.lcLogText + ;
                              PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnFreeTables)), "0") + ")  " + ;
                              UPPER(ALLTRIM(m.lcFreeTable)) +  ;
                              IIF(ISNULL(m.lcDBCXCap) OR EMPTY(m.lcDBCXCap), SPACE(0), " [" + m.lcDBCXCap + "]") + ;
                              SPACE(1) + ccSUCCESSMSG + ;
                              "(" + ;
                              LOWER(FULLPATH(DBF(m.lcFreeTable))) + ;
                              ")" 

               * Get the date/time for the table.
               ldFileTimeStamp = FDATE(DBF(m.lcFreeTable), 1)
               lcFileTimeStamp = TRANSFORM(m.ldFileTimeStamp)

               lcLogText = m.lcLogText + ;
                           " {" + m.lcFileTimeStamp + "} "

               * See if there are issues with tags missing
               * First, see what is live in the table
               lnTags    = ATAGINFO(laTags, SPACE(0), lcFreeTable)

               lcLogText = m.lcLogText + ;
                           ":  " + ;
                           TRANSFORM(RECCOUNT(lcFreeTable)) + SPACE(1) + ccRECORDSLITERAL + ", " + ;
                           TRANSFORM(lnTags) + SPACE(1) + ccTAGSLITERAL

               lcLogText = m.lcLogText + ;
                           IIF(RECCOUNT(lcFreeTable) = 0, ccSPACEGAP + ccEMPTYTABLELITERAL, "")

               * Second, see how many tags are in the Stonefield Database Toolkit / DBCX metadata
               lnCoreMetaTags = GetTypeCount(SPACE(0), LOWER(m.lcFreeAlias) + ".", "I")
               
               IF ISNULL(m.lnCoreMetaTags)
                  * Nothing to report for the index tags because no DBCX metadata
               ELSE
                  IF m.lnTags # m.lnCoreMetaTags
                     lcLogText = m.lcLogText + ;
                                 ccCRLF + ccSPACEGAP + ; 
                                 "- " + ccBADINDEXCOUNTLITERAL + " - " + ;
                                 TRANSFORM(m.lnTags) + " table tags and " + ;
                                 TRANSFORM(m.lnCoreMetaTags) + " metadata tags"
                  ENDIF 
               ENDIF 

               * Display file size details and determine if there are issues with file sizes
               DO CASE
                  * Only show if there are problems with the file size getting to percentage of limit
                  CASE UPPER(ccFILESPECSTYLE) = "SHORT"              
                     lcLogText = m.lcLogText + ;
                                 FileSpecificationsShort(m.lcFreeTable, .T.)

                  * Show regardless if there are problems
                  CASE UPPER(ccFILESPECSTYLE) = "LONG"
                     lcLogText = m.lcLogText + ;
                                 FileSpecificationsLong(m.lcFreeTable)

                  OTHERWISE
                     * No file specifications
                     
               ENDCASE

               * Close this table and move on to the next one
               USE IN (SELECT(m.lcFreeTable))

               lcLogText = m.lcLogText + ccCRLF

               *? Potential optional future enhancements for the process:
               *?  - Scan tables
               *?  - Read memo
               
            CATCH TO loException   
               lnFailedFree  = m.lnFailedFree + 1 
               lcLogText     = m.lcLogText + ;
                               ccCRLF + ;
                               PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnFreeTables)), "0") + ")  " + ;
                               ALLTRIM(m.lcFreeTable) + SPACE(1) + ccFAILEDMSG + ;
                               m.loException.Message + ;
                               " on line " + TRANSFORM(m.loException.LineNo) + ;
                               ccCRLF + ccCRLF

            ENDTRY

            lnI = m.lnI + 1 
            SKIP +1 IN curFreeTables
         ENDDO
      ELSE
         * No DBC selected, nothing to do.
      ENDIF 
      
      WAIT CLEAR 
      
      lcLogText = m.lcLogText + ;
                  ccCRLF + ;
                  TRANSFORM(m.lnOpened) + " DBC table(s) opened successfully" + ccCRLF + ;
                  TRANSFORM(m.lnFailed) + " DBC table(s) failed miserably" + ccCRLF + ;
                  TRANSFORM(m.lnOpenedFree) + " free table(s) opened successfully" + ccCRLF + ;
                  TRANSFORM(m.lnFailedFree) + " free table(s) failed miserably" + ccCRLF + ;
                  ccCRLF + ;
                  ccTOOLNAME + " is complete: " + TRANSFORM(DATETIME())
      
      * RAS 02-Aug-2020, Added the datetime stamp to log file
      lcLogFile = FORCEEXT(ccLOGFILEBASE + "_" + TTOC(DATETIME(), 1), "txt")
      
      STRTOFILE(m.lcLogText, m.lcLogFile, 0)
      MODIFY FILE (m.lcLogFile) NOEDIT RANGE 1,1 NOWAIT 

      CLOSE DATABASES 
   ENDIF 
   
CATCH TO loException
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(m.lcCode, ;
              0+48, ;
              _screen.Caption)
 
ENDTRY

SET SAFETY &lcOldSafety
SET TABLEPROMPT &lcOldTablePrompt 


CLOSE DATABASES ALL 

RETURN  


********************************************************************************
*  METHOD NAME: CollectDBCXMetadata
*
*  AUTHOR: Richard A. Schummer, February 2016
*
*  METHOD DESCRIPTION:
*    This method is called to build a temp cursor of DBCX metadata in case 
*    CoreMeta is used in the data integrity checking process.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CollectDBCXMetadata()

LOCAL llReturnVal, ;
      lnOldSelect

llReturnVal = .T.

TRY 
   lnOldSelect = SELECT()
   
   USE CoreMeta IN 0 SHARED NOUPDATE 
   
   SELECT * ;
      FROM CoreMeta ;
      INTO TABLE (ADDBS(SYS(2023)) + "curDBCXCoreMeta") NOFILTER 

   SELECT curDBCXCoreMeta
   INDEX ON cDBCName   TAG DBCName ADDITIVE 
   INDEX ON cObjectNam TAG ObjName ADDITIVE 
   INDEX ON cRecType   TAG RecType ADDITIVE 
   
   USE IN (SELECT("CoreMeta"))
   
   SELECT (lnOldSelect)

CATCH TO loException
   llReturnVal = .F.
   
ENDTRY

RETURN m.llReturnVal


********************************************************************************
*  METHOD NAME: GetDBCXMetaCaption
*
*  AUTHOR: Richard A. Schummer, May 2018
*
*  METHOD DESCRIPTION:
*    This method returns the DBCX Caption for the table.
*
*  INPUT PARAMETERS:
*    tcDatabase   = required, character, can be left empty for free tables, this
*                   is the database name.
*    tcObjectName = required, character, cannot be empty, name of the free table
*                   to get the DBCX caption.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE GetDBCXMetaCaption(tcDatabase, tcObjectName)

IF PCOUNT() = 2
   IF FILE("coremeta.DBF")
      tcDatabase   = LOWER(JUSTSTEM(m.tcDatabase))
      tcObjectName = LOWER(m.tcObjectName)
   
      SELECT * ;
         FROM coremeta ;
         WHERE cDBCName = m.tcDatabase ;
           AND cObjectNam = m.tcObjectName ;
           AND cRecType = "T" ;
         ORDER BY cObjectNam ;
         INTO CURSOR curMetaCaption
      
      IF _tally > 0
         lcReturnVal = curMetaCaption.cCaption
      ELSE
         lcReturnVal = NULL
      ENDIF 
   ELSE
      lcReturnVal = NULL
   ENDIF 
ELSE
   lcReturnVal = NULL
ENDIF 

RETURN m.lcReturnVal


********************************************************************************
*  METHOD NAME: GetTypeCount
*
*  AUTHOR: Richard A. Schummer, February 2016
*
*  METHOD DESCRIPTION:
*    This method returns the count of detail records found in the SDT Metadata.
*    This could be something like the number of tables registered in SDT, or the 
*    number of indexes in for a table, or the number of fields in a table.
*
*  INPUT PARAMETERS:
*    tcDBCName   = character, required, name of the DBC queried in SDT Metadata
*    tcObjectNam = character, required, name of the object queried in SDT Metadata 
*    tcRecType   = character, required, name of the record type queried in SDT Metadata
* 
*  OUTPUT PARAMETERS:
*    luReturnVal = normally an integer, but NULL if something is not there to be
*                  counted, especially in the case where CoreMeta does not exist.
* 
********************************************************************************
PROCEDURE GetTypeCount(tcDBCName, tcObjectNam, tcRecType)

LOCAL luReturnVal, ;
      lnOldSelect

luReturnVal = NULL

TRY 
   lnOldSelect = SELECT()

   SELECT * ;
      FROM curDBCXCoreMeta ;
      WHERE cDBCName = tcDBCName ;
        AND cObjectNam = tcObjectNam ;
        AND cRecType = tcRecType ;
        AND NOT DELETED() ;
      INTO CURSOR curObjects 

   luReturnVal = RECCOUNT("curObjects")

   USE IN (SELECT("curObjects"))
   
   SELECT (lnOldSelect)


CATCH TO loException
   luReturnVal = NULL

ENDTRY

RETURN m.luReturnVal



********************************************************************************
*  METHOD NAME: FileSpecificationsLong
*
*  AUTHOR: Richard A. Schummer, February 2015
*
*  METHOD DESCRIPTION:
*    This method collects information about the DBF, CDX, and FPT as well as 
*    details about a file size warning if files are closing in on 2GB limit.
*
*  INPUT PARAMETERS:
*    tcAlias = character, required, alias of table to be analyzed.
*    tlWarningOnly = logical, true means include warning, otherwise skip it.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE FileSpecificationsLong(tcAlias, tlFileSizeWarningOnly)

LOCAL lcLogText, ;
      lcOldCompatible, ;
      lnDBFSize, ;
      lnCDXSize, ;
      lnFPTSize, ;
      loException

lcLogText = SPACE(0)

* Record file sizes
lcOldCompatible = SET("Compatible")
SET COMPATIBLE ON 

lnDBFSize = FSIZE(FORCEEXT(DBF(m.tcAlias), "dbf"))

IF m.tlFileSizeWarningOnly
   lcLogText = m.lcLogText + ccSPACEGAP + ;
               "- DBF is " + ;
               FileSizeWarningCheck(m.lnDBFSize) + ;
               ccCRLF
ELSE 
   lcLogText = m.lcLogText + ccSPACEGAP + ;
               "- DBF File Size: " + ;
               ALLTRIM(TRANSFORM(m.lnDBFSize, "999,999,999,999")) + ;
               SPACE(1) + ;
               FileSizeWarningCheck(m.lnDBFSize) + ;
               ccCRLF
ENDIF 

TRY 
   lnCDXSize = FSIZE(FORCEEXT(DBF(m.tcAlias), "cdx"))

   IF m.tlFileSizeWarningOnly
      lcLogText = m.lcLogText + ccSPACEGAP + ;
                  "- CDX is " + ;
                  FileSizeWarningCheck(m.lnDBFSize) + ;
                  ccCRLF
   ELSE 
      lcLogText = m.lcLogText + ccSPACEGAP + ;
                  "- CDX File Size: " + ;
                  ALLTRIM(TRANSFORM(m.lnCDXSize, "999,999,999,999")) + ;
                  SPACE(1) + ;
                  FileSizeWarningCheck(m.lnCDXSize) + ;
                  ccCRLF
   ENDIF 
   
CATCH TO loException
   * Ignore, there is nothing to do if there is no file available
   
ENDTRY

TRY
   lnFPTSize = FSIZE(FORCEEXT(DBF(m.tcAlias), "fpt")) 

   IF m.tlFileSizeWarningOnly
      lcLogText = m.lcLogText + ccSPACEGAP + ;
                  "- FPT is " + ;
                  FileSizeWarningCheck(m.lnDBFSize) + ;
                  ccCRLF
   ELSE 
      lcLogText = m.lcLogText + ccSPACEGAP + ;
                  "- FPT File Size: " + ;
                  ALLTRIM(TRANSFORM(m.lnFPTSize , "999,999,999,999")) + ;
                  SPACE(1) + ;
                  FileSizeWarningCheck(m.lnFPTSize) + ;
                  ccCRLF
   ENDIF 
   
CATCH TO loException      
   * Ignore, there is nothing to do if there is no file available
   
ENDTRY

SET COMPATIBLE &lcOldCompatible


RETURN m.lcLogText


********************************************************************************
*  METHOD NAME: FileSpecificationsShort
*
*  AUTHOR: Richard A. Schummer, February 2018
*
*  METHOD DESCRIPTION:
*    This method collects information about the DBF, CDX, and FPT as well as 
*    details about a file size warning if files are closing in on 2GB limit.
*
*  INPUT PARAMETERS:
*    tcAlias         = character, required, alias of table to be analyzed.
*    tlOnlyIfProblem = logical, true means include warning only if reported, 
*                      otherwise skip it.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE FileSpecificationsShort(tcAlias, tlOnlyIfProblem)

LOCAL lcLogText, ;
      lcOldCompatible, ;
      lnDBFSize, ;
      lnCDXSize, ;
      lnFPTSize, ;
      loException

lcLogText = SPACE(0)

* Record file sizes
lcOldCompatible = SET("Compatible")
SET COMPATIBLE ON 

lnDBFSize         = FSIZE(FORCEEXT(DBF(m.tcAlias), "dbf"))
lcFileSizeWarning = FileSizeWarningCheck(m.lnDBFSize)

IF EMPTY(lcFileSizeWarning) AND m.tlOnlyIfProblem 
   * Nothing to print
ELSE 
   IF NOT EMPTY(lcFileSizeWarning)
      lcLogText = m.lcLogText + ;
                  ccCRLF + ccSPACEGAP + ;
                  "- DBF: " + ;
                  ALLTRIM(TRANSFORM(m.lnDBFSize, "999,999,999,999")) + ;
                  SPACE(1) + ;
                  lcFileSizeWarning
   ENDIF 
ENDIF 

TRY 
   lnCDXSize         = FSIZE(FORCEEXT(DBF(m.tcAlias), "cdx"))
   lcFileSizeWarning = FileSizeWarningCheck(m.lnCDXSize)

   IF EMPTY(lcFileSizeWarning) AND m.tlOnlyIfProblem 
      * Nothing to print
   ELSE 
      IF NOT EMPTY(lcFileSizeWarning)
         lcLogText = m.lcLogText + ;
                     ccCRLF + ccSPACEGAP + ;
                     "- CDX: " + ;
                     ALLTRIM(TRANSFORM(m.lnCDXSize, "999,999,999,999")) + ;
                     SPACE(1) + ;
                     lcFileSizeWarning
      ENDIF 
   ENDIF
      
CATCH TO loException
   * Ignore, there is nothing to do if there is no file available
   
ENDTRY

TRY
   lnFPTSize = FSIZE(FORCEEXT(DBF(m.tcAlias), "fpt")) 
   lcFileSizeWarning = FileSizeWarningCheck(m.lnFPTSize)

   IF EMPTY(lcFileSizeWarning) AND m.tlOnlyIfProblem 
      * Nothing to print
   ELSE 
      IF NOT EMPTY(lcFileSizeWarning)
         lcLogText = m.lcLogText + ;
                     ccCRLF + ccSPACEGAP + ;
                     "- FPT: " + ;
                     ALLTRIM(TRANSFORM(m.lnFPTSize , "999,999,999,999")) + ;
                     SPACE(1) + ;
                     lcFileSizeWarning
                     
      ENDIF 
   ENDIF 
   
CATCH TO loException      
   * Ignore, there is nothing to do if there is no file available
   
ENDTRY

SET COMPATIBLE &lcOldCompatible

RETURN m.lcLogText


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
PROCEDURE FileSizeWarningCheck(tnFileSize, tcOption)

LOCAL lcFileSizeStatus, ;
      lnFileLimitWarningLevel, ;
      lnPercentage

IF PCOUNT() < 2
   tcOption = SPACE(0)
ENDIF 

tnFilesize    = m.tnFilesize * 1.0000000000
lnTwoGigLimit = cnTWOGIGLIMIT

lnOldDecimals = SET("Decimals")

SET DECIMALS TO 10

IF VARTYPE(tnFileSize) = "N"
   * The reason for dividing each value by 1024 is to get the math away from the 2GB limit for integers.
   lnFileLimitWarningLevel = (ROUND((cnTWOGIGLIMIT/1024) * (cnHOWCLOSETOLIMIT/1024), 0)) * 1024
   lnPercentage            = ROUND((m.tnFileSize/1024) / (m.lnTwoGigLimit/1024), 10) * 100
 
   IF m.tcOption = "FULL"
      lcFileSizeStatus = ALLTRIM(TRANSFORM(m.lnPercentage, "999.99")) + "% of VFP limit"
   ELSE
      lcFileSizeStatus = SPACE(0)
   ENDIF 
   
   IF  m.lnFileLimitWarningLevel > m.tnFileSize
      * All is good in the world, nothing else to report
   ELSE
      * Raise a flag that file is getting close to VFP limits
      lcFileSizeStatus = m.lcFileSizeStatus +  ;
                         "WARNING: getting close to file size limit (" + ALLTRIM(TRANSFORM(TRANSFORM(cnHOWCLOSETOLIMIT, "999.99"))) + "%)"
   ENDIF

   IF NOT EMPTY(m.lcFileSizeStatus)
      lcFileSizeStatus = "(" + m.lcFileSizeStatus + ")"
   ENDIF 
ELSE
   lcFileSizeStatus = "Valid file size not passed to checker"
ENDIF 

SET DECIMALS TO (m.lnOldDecimals)

RETURN ALLTRIM(m.lcFileSizeStatus)


********************************************************************************
*  METHOD NAME: CloseData
*
*  AUTHOR: Richard A. Schummer, February 2016
*
*  METHOD DESCRIPTION:
*    This method is called to close any open cursors for the process.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CloseData()

USE IN (SELECT("curDBCXCoreMeta"))
USE IN (SELECT("curFreetables"))
USE IN (SELECT("curMetaCaption"))

RETURN 

*: EOF :*  