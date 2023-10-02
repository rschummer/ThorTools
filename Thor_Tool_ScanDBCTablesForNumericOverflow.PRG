******************************************************************************************
*  PROGRAM: Thor_Tool_ScanDBCTablesForNumericOverflow.prg
*
*  AUTHOR: Richard A. Schummer, December 2014
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
*     This program checks to see if the data contains a numeric overflow condition for 
*     all tables in a DBC, or a folder.
*
*  CALLING SYNTAX:
*     DO Thor_Tool_ScanDBCTablesForNumericOverflow.prg
*     Thor_Tool_ScanDBCTablesForNumericOverflow()
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
* 12/08/2014  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     1.1      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
* 09/21/2015  Richard A. Schummer     1.2      Alternative way of checking instead of stars. 
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

IF PCOUNT() = 1 AND 'O' = VARTYPE(m.lxParam1) AND 'thorinfo' == LOWER(m.lxParam1.Class)
   WITH m.lxParam1
      * Required
      .Prompt         = 'Find Numeric Overflow in DBFs of DBC' && used in menus
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2 && a description for the tool
         This program scans through all tables in a Database Container, folder, or DBCX metadata to find numeric overflow conditions in the table's numeric columns. 
      ENDTEXT  
      
      .StatusBarText   = 'Scan through a FoxPro table and look for numeric overflow stars in numeric columns'
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source        = "WLC"                      && where did this tool come from?  Your own initials, for instance
      .Category      = "WLC"                      && creates categorization of tools; defaults to .Source if empty
      .Sort          = 0                          && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version       = "Version 1.1, September 20, 2015"           && e.g., 'Version 7, May 18, 2011'
      .Author        = "Rick Schummer"
      .Link          = "https://github.com/rschummer/ThorTools"    && link to a page for this tool
      .VideoLink     = SPACE(0)                                    && link to a video for this tool
      
   ENDWITH 

   RETURN m.lxParam1
ENDIF 

IF PCOUNT() = 0
   DO ToolCode
ELSE
   DO ToolCode WITH m.lxParam1
ENDIF

RETURN

********************************************************************************
*  METHOD NAME: ToolCode
*
*  AUTHOR: Richard A. Schummer, December 2014
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

#DEFINE ccCRLF      CHR(13)+CHR(10)
#DEFINE ccLOGFILE   "DBCNumericOverflowCheckLog.txt"
#DEFINE ccTOOL      "WLC Numeric Overflow Locator - DBC"

LOCAL lcOldSafety, ;
      lcLogText, ;
      lnOpened, ;
      lnFailed, ;
      lnTableProblem, ;
      lnDBCFiles, ;
      laDatabases[1], ;
      lcDBCFile, ;
      lnTables, ;
      laTables[1], ;
      lnJ, ;
      lcFile, ;
      lnFails, ;
      lnFields, ;
      lnI, ;
      lnRows, ;
      laFields[1], ;
      lcNumericFields, ;
      luValue, ;
      lnLenWholePart, ;
      lnLenDecimalPart, ;
      lnLenFullSize, ;
      lnMaxShouldFitValue, ;
      lnMaxVFPJamsInValue, ;
      loException, ;
      lcCode

TRY
   CLEAR 

   lcOldSafety = SET("Safety")
   SET SAFETY OFF

   CD ?

   lcLogText = TRANSFORM(DATETIME()) + ccCRLF

   lnOpened       = 0
   lnFailed       = 0
   lnTableProblem = 0
   lnDBCFiles     = ADIR(laDatabases, "*.DBC")
   
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
                  ccTOOL + ccCRLF + ;
                  "DBC: " + LOWER(FULLPATH(DBC())) + ccCRLF + ;
                  "Current Folder: " + FULLPATH(CURDIR()) + ccCRLF + ccCRLF + ;
                  "Search for 'ERROR' to find the overflow conditions" + ccCRLF + ccCRLF
      
      WAIT WINDOW "Determining number of tables to analyze" NOWAIT            
      lnTables  = ADBOBJECTS(laTables, "TABLE")

      * RAS 07-Mar-2023, sort the tables in alpha order, ascending.
      ASORT(laTables, 1, m.lnTables, 0, 0)
      
      FOR lnJ = 1 TO m.lnTables
         lcFile = laTables[lnJ]
         
         IF EMPTY(m.lcFile)
            * Nothing to do
         ELSE
            TRY 
               lcOldSafety = SET("Safety")
               SET SAFETY OFF

               lnFails  = 0
               lnFields = 0
               lnI      = 0
               lnRows   = 0

               USE IN (SELECT("curChkOverflow"))
               USE (m.lcFile) IN 0 SHARED AGAIN ALIAS curChkOverflow
               SELECT curChkOverflow
               
               lnFields  = AFIELDS(laFields, "curChkOverflow")
               lnOpened  = m.lnOpened + 1 
               lcLogText = m.lcLogText + ;
                           PADL(TRANSFORM(m.lnJ), LENC(TRANSFORM(m.lnTables)), "0") + ")  " + ;
                           ALLTRIM(laTables[lnJ]) + " successfully opened..." + ;
                           " with " + TRANSFORM(RECCOUNT("curChkOverflow")) + " rows - " + ;
                           FULLPATH(DBF("curChkOverflow")) + ;
                           ccCRLF

               WAIT WINDOW "Reviewing " + JUSTSTEM(m.lcFile) + ;
                           " with " + TRANSFORM(RECCOUNT("curChkOverflow")) + " rows - " + ;
                           "...(" + TRANSFORM(lnJ) + " of " + TRANSFORM(lnTables) + ")" ;
                    NOWAIT NOCLEAR 

               lcNumericFields = SPACE(0)

               FOR m.lnI = 1 TO m.lnFields
                  m.luValue = EVALUATE(FIELD(m.lnI, "curChkOverflow"))
               
                  IF VARTYPE(m.luValue) = "N"
                     m.lcNumericFields = m.lcNumericFields + ;
                                         TRANSFORM(lnI, "9999") + ")" + SPACE(1) + ;
                                         LOWER(FIELD(m.lnI, "curChkOverflow"))  + SPACE(1) + ;
                                         "-" + SPACE(1) + ;
                                         "Type:" + SPACE(1) + DataTypeTranslator(UPPER(laFields[lnI,2])) + SPACE(1) + ;
                                         "Size:" + SPACE(1) + TRANSFORM(laFields[lnI,3]) + "," + TRANSFORM(laFields[lnI,4]) + ;
                                         ccCRLF
                  ENDIF 
               ENDFOR 
               
               lnFailed  = 0

               IF EMPTY(m.lcNumericFields)
                  * Nothing to scan, note it in the log.
                  lcLogText = m.lcLogText + "NO Numeric Fields" + ccCRLF
               ELSE 
                  lcLogText = m.lcLogText + "Numeric Fields: " + ccCRLF + m.lcNumericFields + ccCRLF
                  lnRows    = 0 

                  SCAN        
                     lnRows = m.lnRows + 1 
                     
                     FOR m.lnI = 1 TO m.lnFields
                        m.luValue = EVALUATE(FIELD(m.lnI, "curChkOverflow"))
                     
                        * NOTE: Currency does not evaluate to numeric, but VFP enforces 
                        *       values in proper range with error 1988 (Currency value is out of range.)
                        IF VARTYPE(m.luValue) = "N"
                           * Determine the maximum value that should fit properly
                           lnLenWholePart      = laFields[m.lnI,3] - laFields[m.lnI,4] - IIF(laFields[m.lnI,4] > 0, 1, 0)
                           lnLenDecimalPart    = laFields[m.lnI,4]
                           lnLenFullSize       = laFields[m.lnI,3]
                           
                           * Check if Integer with 4,0 definition
                           IF laFields[m.lnI,2] = "I" AND laFields[m.lnI,3] = 4 AND laFields[m.lnI,4] = 0
                              * Integer columns have fixed limit
                              lnMaxPosShouldFitValue = 2147483647
                              lnMaxPosVFPJamsInValue = m.lnMaxPosShouldFitValue
                              lnMaxNegShouldFitValue = -2147483647
                              lnMaxNegVFPJamsInValue = m.lnMaxNegShouldFitValue
                           ELSE 
                              * Rest have column size limit set by developer
                              lnMaxPosShouldFitValue = VAL(REPLICATE("9", m.lnLenWholePart) + "." + REPLICATE("9", m.lnLenDecimalPart))
                              lnMaxPosVFPJamsInValue = VAL(REPLICATE("9", m.lnLenFullSize))
                              lnMaxNegShouldFitValue = VAL(REPLICATE("9", m.lnLenWholePart -1) + "." + REPLICATE("9", m.lnLenDecimalPart)) * (-1)
                              lnMaxNegVFPJamsInValue = VAL(REPLICATE("9", m.lnLenFullSize - 1)) * (-1)
                           ENDIF 
                           
                           * First check if hard core "stars" VFP overflow condition that throws errors
                           IF m.luValue > m.lnMaxPosVFPJamsInValue OR m.luValue < m.lnMaxNegVFPJamsInValue
                              lnFailed  = m.lnFailed + 1 
                              lcLogText = m.lcLogText + ;
                                          "ERROR: VFP Numeric Overflow on record " + ;
                                          TRANSFORM(RECNO("curChkOverflow")) + ;
                                          " in field: " + ;
                                          LOWER(FIELD(m.lnI, "curChkOverflow")) + ;
                                          ccCRLF
                           ELSE 
                              * Secondly, check if the data exceeds the natural design of the column.
                              IF m.luValue > m.lnMaxPosShouldFitValue OR m.luValue < m.lnMaxNegShouldFitValue
                                 lnFailed  = m.lnFailed + 1 
                                 lcLogText = m.lcLogText + ;
                                             "ERROR: Column Design Numeric Overflow on record " + ;
                                             TRANSFORM(RECNO("curChkOverflow")) + ;
                                             " in field: " + ;
                                             LOWER(FIELD(m.lnI, "curChkOverflow")) + ;
                                             ccCRLF
                              ENDIF 
                           ENDIF
                        ENDIF 
                     ENDFOR 
                  ENDSCAN
               ENDIF 

               lcLogText = m.lcLogText + ;
                           "Overflow conditions discovered: " + TRANSFORM(m.lnFailed) + REPLICATE(ccCRLF, 3)
               
               lnTableProblem = m.lnTableProblem + IIF(m.lnFailed > 0, 1, 0)
               
               USE IN (SELECT("curChkOverflow"))
               
            CATCH TO m.loException 
               IF m.loException.ErrorNo = 1
                  lcCode = "Error: " + m.loException.Message + ;
                     " [" + TRANSFORM(m.loException.Details) + "] " 

                  MESSAGEBOX(m.lcCode, ;
                             0+48, ;
                             _screen.Caption)
               ELSE 
                  lcCode = "Error: " + m.loException.Message + ;
                     " [" + TRANSFORM(m.loException.Details) + "] " + ;
                     " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
                     " in " + m.loException.Procedure + ;
                     " on " + TRANSFORM(m.loException.LineNo)

                  MESSAGEBOX(m.lcCode, ;
                             0+48, ;
                             _screen.Caption)
               ENDIF  
            ENDTRY
         ENDIF 
      ENDFOR 
      
      WAIT CLEAR 
   ENDIF 

   lcLogText = m.lcLogText + ;
               ccCRLF + ;
               TRANSFORM(m.lnOpened) + " tables opened successfully" + ccCRLF + ;
               TRANSFORM(m.lnTableProblem) + " tables have numeric overflow problem" + ccCRLF + ccCRLF + ;
               ccTOOL + " is complete: " + TRANSFORM(DATETIME())
   
   STRTOFILE(lcLogText, ccLOGFILE, 0)
   MODIFY FILE (ccLOGFILE) NOEDIT NOWAIT RANGE 1,1
   
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
*  METHOD NAME: DataTypeTranslator
*
*  AUTHOR: Richard A. Schummer, September 2015
*
*  METHOD DESCRIPTION:
*    Translates the data type letter codes to natural language data type.
*
*  INPUT PARAMETERS:
*    tcDataType = character, required, single letter data type
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE DataTypeTranslator(tcDataType)

LOCAL lcReturnVal

IF VARTYPE(tcDataType) # "C"
   lcReturnVal = "Bad parameter"
ELSE
   tcDataType = UPPER(tcDataType)

   DO CASE
      CASE tcDataType = "I"
         lcReturnVal = "Integer"

      CASE tcDataType = "N"
         lcReturnVal = "Numeric"

      CASE tcDataType = "Y"
         lcReturnVal = "Currency"

      CASE tcDataType = "B"
         lcReturnVal = "Double"

      CASE tcDataType = "F"
         lcReturnVal = "Float"

      CASE tcDataType = "L"
         lcReturnVal = "Logical"

      CASE tcDataType = "T"
         lcReturnVal = "Data/Time"

      CASE tcDataType = "D"
         lcReturnVal = "Date"

      CASE tcDataType = "C"
         lcReturnVal = "Character"

      CASE tcDataType = "M"
         lcReturnVal = "Memo"

      CASE tcDataType = "W"
         lcReturnVal = "Blob"

      CASE tcDataType = "V"
         lcReturnVal = "Varchar"

      CASE tcDataType = "Q"
         lcReturnVal = "Varbinary"

      CASE tcDataType = "G"
         lcReturnVal = "General"

      OTHERWISE
         lcReturnVal = ""

   ENDCASE
ENDIF 

RETURN m.lcReturnVal


*: EOF :*  