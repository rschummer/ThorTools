******************************************************************************************
*  PROGRAM: Thor_Tool_CheckDataIntegrity.prg
*
*  AUTHOR: Richard A. Schummer, November 2013
*
*  COPYRIGHT � 2013-2015   All Rights Reserved.
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
*
******************************************************************************************
LPARAMETERS lxParam1

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
      .Version       = "Version 2.2, September 20, 2015"           && e.g., 'Version 7, May 18, 2011'
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

#DEFINE ccCRLF            CHR(13)+CHR(10)
#DEFINE ccLOGFILE         "DBCIntegrityCheckLog.txt"
#DEFINE ccDBCVALIDATEFILE "DBCValidate.txt"

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
      lnTables

lcOldSafety = SET("Safety")
SET SAFETY OFF

CD ?

lcLogText = TRANSFORM(DATETIME()) + ccCRLF

TRY
   lnDBCFiles = ADIR(laDatabases, "*.DBC")
   
   IF lnDBCFiles = 1
      lcDBCFile = laDatabases[1, 1]
   ELSE
      lcDBCFile = GETFILE("dbc")
   ENDIF 
 
   IF EMPTY(lcDBCFile)
      * Nothing to do
   ELSE
      WAIT WINDOW "Opening database, please wait..." NOWAIT 
      
      OPEN DATABASE (m.lcDBCFile) SHARED 
      SET DATABASE TO (JUSTSTEM(m.lcDBCFile))

      lcLogText = m.lcLogText + ;
                  "DBC: " + LOWER(FULLPATH(DBC())) + ccCRLF + ;
                  "Current Folder: " + FULLPATH(CURDIR()) + ccCRLF + ccCRLF
      
      WAIT WINDOW "Determining number of tables to analyze" NOWAIT            
      lnTables  = ADBOBJECTS(laTables, "TABLE")
      
      lcLogText = m.lcLogText + ;
                  TRANSFORM(lnTables) + " Tables to check." + ;
                  ccCRLF 
      
      WAIT WINDOW "Validating database, please wait..." NOWAIT
      
      VALIDATE DATABASE NOCONSOLE TO FILE (ccDBCVALIDATEFILE)
      
      lcLogText = m.lcLogText + ;
                  FILETOSTR(ccDBCVALIDATEFILE) + ;
                  ccCRLF + ccCRLF
      
      DELETE FILE (ccDBCVALIDATEFILE) RECYCLE 

      lnOpened = 0
      lnFailed = 0
      
      IF lnTables > 0
         =ASORT(laTables)
         
         lnI = 1

         DO WHILE m.lnI <= m.lnTables 
         
            TRY
               WAIT WINDOW LOWER(laTables[lnI]) + " - " + TRANSFORM((m.lnI/m.lnTables)*100) + "%" NOWAIT 
               
               USE (laTables[lnI]) SHARED AGAIN IN 0 NOUPDATE 
               lnOpened  = m.lnOpened + 1 
               lcLogText = m.lcLogText + ;
                           PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnTables)), "0") + ")  " + ;
                           ALLTRIM(laTables[lnI]) + " successfully opened..." + ;
                           FULLPATH(DBF(laTables[lnI])) + ccCRLF

               USE IN (SELECT(laTables[lnI]))
               
               *? Potential optional future enhancements for the process:
               *?  - Scan tables
               *?  - Read memo
               
            CATCH TO loException   
               lnFailed  = m.lnFailed + 1 
               lcLogText = m.lcLogText + ;
                           ccCRLF + ;
                           PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnTables)), "0") + ")  " + ;
                           ALLTRIM(laTables[lnI]) + " ************** FAILED..." + m.loException.Message + ;
                           ccCRLF + ccCRLF

            ENDTRY

            lnI = m.lnI + 1 
         ENDDO
      ENDIF 
      
      * Check free tables.
      lnOpenedFree = 0
      lnFailedFree = 0

      IF FILE("coremeta.DBF")
         SELECT * ;
            FROM coremeta ;
            WHERE EMPTY(cDBCName) ;
              AND cRecType = "T" ;
            INTO CURSOR curFreetables
         
         IF USED("curFreetables")
            lnFreeTables = RECCOUNT("curFreetables")

            lcLogText    = m.lcLogText + ccCRLF + ;
                           "FREE Tables..." + ccCRLF
         ELSE
            lnFreeTables = 0
         ENDIF 
         
         lnI = 1
         
         DO WHILE m.lnI <= m.lnFreeTables 
            TRY
               lcFreeTable = ALLTRIM(curFreetables.cObjectNam)
               
               WAIT WINDOW LOWER(m.lcFreeTable) + " - " + TRANSFORM((m.lnI/m.lnFreeTables)*100) + "%" NOWAIT 
               
               USE (m.lcFreeTable) SHARED AGAIN IN 0 NOUPDATE 
               lnOpenedFree = m.lnOpenedFree + 1 
               lcLogText    = m.lcLogText + ;
                              PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnFreeTables)), "0") + ")  " + ;
                              ALLTRIM(m.lcFreeTable) + " successfully opened..." + ;
                              FULLPATH(DBF(m.lcFreeTable)) + ccCRLF

               USE IN (SELECT(m.lcFreeTable))
               
               *? Potential optional future enhancements for the process:
               *?  - Scan tables
               *?  - Read memo
               
            CATCH TO loException   
               lnFailedFree  = m.lnFailedFree + 1 
               lcLogText     = m.lcLogText + ;
                               ccCRLF + ;
                               PADL(TRANSFORM(m.lnI), LENC(TRANSFORM(m.lnTables)), "0") + ")  " + ;
                               ALLTRIM(laTables[lnI]) + " ************** FAILED..." + m.loException.Message + ;
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
                  "DBC Integrity Check is complete: " + TRANSFORM(DATETIME())
      
      STRTOFILE(m.lcLogText, ccLOGFILE, 0)
      MODIFY FILE (ccLOGFILE) NOEDIT RANGE 1,1
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

CLOSE DATABASES ALL 

RETURN  

*: EOF :* 