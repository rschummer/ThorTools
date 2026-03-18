******************************************************************************************
*  PROGRAM: Thor_Tool_ScanTableForNumericOverflow.prg
*
*  AUTHOR: Richard A. Schummer, November 2014
*
*  COPYRIGHT © 2014-2026   All Rights Reserved.
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
*     a single table.
*
*  CALLING SYNTAX:
*     DO Thor_Tool_ScanTableForNumericOverflow.prg [ WITH <table name> ]
*     Thor_Tool_ScanTableForNumericOverflow(<table name>)
*
*  INPUT PARAMETERS:
*     lxParam1 = unknown data type, typically the standard Thor object passed to register
*                the tool with Thor and appear on the Thor menu.
*
*                if character, it should be a table name (with or without path).
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
*                                  C H A N G E    L O G                                   
*
*    Date     Developer               Version  Description
* ----------  ----------------------  -------  -------------------------------------------
* 11/01/2014  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     1.1      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
* 03/18/2026  Richard A. Schummer     1.2      Updated to pass in parameter with table name,
*                                              note if record is deleted (now checks is deleted 
*                                              records have issue), and skip all General 
*                                              fields to avoid error during EVAL(), and 
*                                              move log file to temp folder. Added progress
*                                              meter via WAIT WINDOW.
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

IF PCOUNT() = 1 AND 'O' = VARTYPE(m.lxParam1) AND 'thorinfo' == LOWER(m.lxParam1.Class)
   WITH m.lxParam1
      * Required
      .Prompt         = 'Find Numeric Overflow columns in DBF' && used in menus
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2 && a description for the tool
         This program rolls through a specified or selected DBF and finds numeric overflow conditions in the table's numeric columns.      
      ENDTEXT  
      
      .StatusBarText   = 'Scan through a FoxPro table and look for numeric overflow stars in numeric columns.'
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source        = "WLC"          && where did this tool come from?  Your own initials, for instance
      .Category      = "WLC"          && creates categorization of tools; defaults to .Source if empty
      .Sort          = 0              && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version       = "Version 1.2, March 18, 2026"               && e.g., 'Version 7, May 18, 2011'
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
PROCEDURE ToolCode(tcTable)

#DEFINE ccCRLF      CHR(13) + CHR(10)
#DEFINE ccLOGFILE   "NumericOverflowCheckLog.txt"
#DEFINE ccLOGFOLDER ADDBS(SYS(2023))
#DEFINE ccTOOL      "WLC Numeric Overflow Locator - Single Table"

LOCAL lcFile, ;
      lcOldConsole, ;
      lcOldSafety, ;
      lcOldDeleted, ;
      lnFails, ;
      lnFields, ;
      lnI, ;
      lnRows, ;
      lcLogFile, ;
      laFields[1], ;
      lnNumericColCount, ;
      luValue, ;
      lnLenWholePart, ;
      lnLenDecimalPart, ;
      lnLenFullSize, ;
      lnMaxPosShouldFitValue, ;
      lnMaxPosVFPJamsInValue, ;
      lnMaxNegShouldFitValue, ;
      lnMaxNegVFPJamsInValue, ;
      lcMessage, ;
      loException as Exception, ;
      lcCode

CLEAR

IF PCOUNT() = 1 AND VARTYPE(m.tcTable) = "C"
   m.lcFile = FORCEEXT(ALLTRIM(m.tcTable), "dbf")
ELSE
   m.lcFile = GETFILE("dbf")
ENDIF

IF EMPTY(m.lcFile)
   * Nothing to do
ELSE
   TRY
      lcOldConsole = SET("Console")
      SET CONSOLE OFF

      lcOldSafety = SET("Safety")
      SET SAFETY OFF

      lcOldDeleted = SET("Deleted")
      SET DELETED OFF

      lnFails  = 0
      lnFields = 0
      lnI      = 0
      lnRows   = 0

      USE IN (SELECT("curChkOverflow"))
      USE (m.lcFile) IN 0 SHARED AGAIN ALIAS curChkOverflow
      SELECT curChkOverflow

      lcLogFile = ccLOGFOLDER + FORCEEXT(JUSTSTEM(ccLOGFILE) + "_" + LOWER(JUSTSTEM(DBF("curChkOverflow"))), "txt")

      SET ALTERNATE TO (m.lcLogFile)
      SET ALTERNATE ON

      ? ccTOOL
      ? DATETIME()
      ? SPACE(0)
      ? LOWER(DBF("curChkOverflow"))

      lnFields          = AFIELDS(laFields, "curChkOverflow")
      lnNumericColCount = 0

      ? SPACE(0)
      ? "Numeric Columns:"

      FOR lnI = 1 TO lnFields
         * Skip over General fields which throws an error
         IF laFields[lnI, 2] = "G"
            LOOP 
         ELSE
            m.luValue = EVALUATE(FIELD(m.lnI, "curChkOverflow"))
         ENDIF 
         
         IF VARTYPE(m.luValue) = "N"
            ? TRANSFORM(lnI, "9999") + ")", LOWER(FIELD(m.lnI, "curChkOverflow")), "-", ;
               "Type:", DataTypeTranslator(UPPER(laFields[lnI, 2])) + ",", ;
               "Size:", TRANSFORM(laFields[lnI, 3]) + "," + TRANSFORM(laFields[lnI, 4])

            lnNumericColCount = m.lnNumericColCount + 1
         ENDIF
      ENDFOR

      IF m.lnNumericColCount = 0
         ? " - This table does not have any numeric columns of any kind."
      ENDIF

      ? SPACE(0)
      
      lnRecordCount = RECCOUNT()

      SCAN
         lnRows = m.lnRows + 1
         
         DisplayProgressMeter(m.lnRows, m.lnRecordCount, 5)

         FOR m.lnI = 1 TO m.lnFields
            * Skip over General fields which throws an error
            IF laFields[lnI, 2] = "G"
               LOOP 
            ELSE
               luValue = EVALUATE(FIELD(m.lnI, "curChkOverflow"))
            ENDIF 

            * NOTE: Currency does not evaluate to numeric, but VFP enforces 
            *       values in proper range with error 1988 (Currency value is out of range.)
            IF VARTYPE(m.luValue) = "N"
               * Determine the maximum value that should fit properly
               lnLenWholePart   = laFields[lnI, 3] - laFields[lnI, 4] - IIF(laFields[lnI, 4] > 0, 1, 0)
               lnLenDecimalPart = laFields[lnI, 4]
               lnLenFullSize    = laFields[lnI, 3]

               * Check if Integer with 4,0 definition
               IF laFields[lnI, 2] = "I" AND laFields[lnI, 3] = 4 AND laFields[lnI, 4] = 0
                  * Integer columns have fixed limit
                  lnMaxPosShouldFitValue = 2147483647
                  lnMaxPosVFPJamsInValue = lnMaxPosShouldFitValue
                  lnMaxNegShouldFitValue = -2147483647
                  lnMaxNegVFPJamsInValue = lnMaxNegShouldFitValue
               ELSE
                  * Rest have column size limit set by developer
                  lnMaxPosShouldFitValue = VAL(REPLICATE("9", m.lnLenWholePart) + "." + REPLICATE("9", m.lnLenDecimalPart))
                  lnMaxPosVFPJamsInValue = VAL(REPLICATE("9", m.lnLenFullSize))
                  lnMaxNegShouldFitValue = VAL(REPLICATE("9", m.lnLenWholePart - 1) + "." + REPLICATE("9", m.lnLenDecimalPart)) * (-1)
                  lnMaxNegVFPJamsInValue = VAL(REPLICATE("9", m.lnLenFullSize - 1)) * (-1)
               ENDIF

               * First check if hard core "stars" VFP overflow condition that throws errors
               IF m.luValue > m.lnMaxPosVFPJamsInValue OR m.luValue < m.lnMaxNegVFPJamsInValue
                  lnFails = m.lnFails + 1
                  ? "VFP Numeric Overflow on record " + TRANSFORM(RECNO("curChkOverflow")) + IIF(DELETED(), " (DELETED record)", SPACE(0)) + " in field: " + LOWER(FIELD(m.lnI, "curChkOverflow")) 
               ELSE
                  * Secondly, check if the data exceeds the natural design of the column.
                  IF m.luValue > lnMaxPosShouldFitValue OR m.luValue < lnMaxNegShouldFitValue
                     lnFails = m.lnFails + 1
                     ? "Column Design Numeric Overflow on record " + TRANSFORM(RECNO("curChkOverflow")) + IIF(DELETED(), " (DELETED record)", SPACE(0)) + " in field: " + LOWER(FIELD(m.lnI, "curChkOverflow"))
                  ENDIF
               ENDIF
            ENDIF
         ENDFOR
      ENDSCAN
      
      WAIT CLEAR 

      USE IN (SELECT("curChkOverflow"))

      lcMessage = "Total records: " + TRANSFORM(m.lnRows) + ccCRLF + ;
         "Total failures found: " + TRANSFORM(m.lnFails)

      ? SPACE(0)
      ? m.lcMessage
      ? SPACE(0)
      ? DATETIME()
      ? SPACE(0)

      SET ALTERNATE OFF
      SET ALTERNATE TO

      MODIFY FILE (m.lcLogFile) NOEDIT NOWAIT RANGE 1, 1

   CATCH TO m.loException
      IF m.loException.ERRORNO = 1
         lcCode = "Error: " + m.loException.MESSAGE + ;
            " [" + TRANSFORM(m.loException.DETAILS) + "] "

         MESSAGEBOX(m.lcCode, ;
            0 + 48, ;
            _SCREEN.CAPTION)
      ELSE
         lcCode = "Error: " + m.loException.MESSAGE + ;
            " [" + TRANSFORM(m.loException.DETAILS) + "] " + ;
            " (" + TRANSFORM(m.loException.ERRORNO) + ")" + ;
            " in " + m.loException.PROCEDURE + ;
            " on " + TRANSFORM(m.loException.LINENO)

         MESSAGEBOX(m.lcCode, ;
            0 + 48, ;
            _SCREEN.CAPTION)
      ENDIF

   FINALLY
      * Clean up, just in case of error
      SET ALTERNATE OFF
      SET ALTERNATE TO

      * Normal clean up
      SET CONSOLE &lcOldConsole
      SET DELETED &lcOldDeleted

   ENDTRY
ENDIF

RETURN


********************************************************************************
*  METHOD NAME: DisplayProgressMeter
*
*  AUTHOR: Richard A. Schummer, March 2026
*
*  METHOD DESCRIPTION:
*    tnRows = (required), integer, number of rows processed to this point in time.
*    tnRecordCount =(required) integer, number of rows to be processed.
*    tnEveryPercentageToReportProgress = integer, defaults to 10, increments of 
*                                        percentage to display the progress.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE DisplayProgressMeter(tnRows, tnRecordCount, tnEveryPercentageToReportProgress)

LOCAL nEveryPercentageToReportProgress, ;
      lnPercent, ;
      lnPercentRounded
      
IF PCOUNT() < 3
   lnEveryPercentageToReportProgress = 10
ELSE 
   lnEveryPercentageToReportProgress = m.tnEveryPercentageToReportProgress
ENDIF 

lnPercent        = (m.tnRows / m.tnRecordCount) * 100
lnPercentRounded = ROUND(m.lnPercent, 0)

IF MOD(m.lnPercentRounded, m.lnEveryPercentageToReportProgress) = 0
   WAIT WINDOW REPLICATE("|", m.lnPercentRounded) + " (" + TRANSFORM(m.lnPercentRounded) + "%)" NOWAIT NOCLEAR
ENDIF

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

IF VARTYPE(m.tcDataType) # "C"
   lcReturnVal = "Bad parameter"
ELSE
   tcDataType = UPPER(m.tcDataType)

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