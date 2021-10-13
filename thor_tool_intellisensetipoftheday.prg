******************************************************************************************
*  PROGRAM: Thor_Tool_IntelliSenseTipOfTheDay.prg
*
*  AUTHOR: Richard A. Schummer, May 2021
*
*  COPYRIGHT © 2021   All Rights Reserved.
*     White Light Computing, Inc.
*     PO Box 391
*     Washington Twp., MI  48094
*     raschummer@whitelightcomputing.com
*
*  PROGRAM DESCRIPTION:
*    This tool produces a random Intellisense Tip of the Day for all the 
*    Intellisense entries that are not scripts. It is designed to prompt 
*    memory of the custom entry, based on the filter set up in the constant 
*    ccISTOTD_FILTER found in the ToolCode method.
*
*    Tip of the day can be displayed on the VFP desktop, in a message box or 
*    wait window as defined by the constant ccISTOTD_MSGSTYLE.
*
*  CALLING SYNTAX:
*     Just drop this into your Thor MyTools folder to automatically add to menu
*
*     DO Thor_Tool_IntelliSenseTipOfTheDay.prg
*     Thor_Tool_IntelliSenseTipOfTheDay()
*
*  INPUT PARAMETERS:
*     None
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
*                                   C H A N G E    L O G                                  
*
*    Date     Developer               Version  Description
* ----------  ----------------------  -------  -------------------------------------------
* 05/12/2021  Richard A. Schummer     1.0      Created program
* ----------------------------------------------------------------------------------------
* 05/18/2021  Richard A. Schummer     1.1      Minor bug fixes in formatting output, added
*                                              ability to define how many tips are output.
* ----------------------------------------------------------------------------------------
* 05/24/2021  Richard A. Schummer     1.2      Added FoxCode record number to the output, 
*                                              record count of result set now displayed if
*                                              desktop mode, and FoxCode table. Added 
*                                              function parameters for function entries.
*                                              Removed Special types, and blank Abbreviation
* ----------------------------------------------------------------------------------------
* 05/28/2021  Richard A. Schummer     1.3      Removed showing deleted records
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS txParam1

#DEFINE ccCRLF       CHR(13)+CHR(10)

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND 'O' = VARTYPE(txParam1) AND 'thorinfo' == LOWER(txParam1.Class)
   WITH txParam1
   
      * Required and used in menus
      .Prompt          = "Intellisense Tip of the Day"           
      .StatusBarText   = "This tool displays random Intellisense file (FoxCode) entries to establish and reinforce new habits"  
      
      * Optional, a description for the tool, easy to overwrite with addtional text, even after the status bar text.
      TEXT TO .Description TEXTMERGE NOSHOW PRETEXT 1+2 
         <<.StatusBarText>>      
      ENDTEXT  
      
      * Determine if you want to be able to run when Thor starts up (enables checkbox if you want to run it).
      .CanRunAtStartUp = .T.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Category      = "WLC"                      && creates categorization of tools; defaults to .Source if empty
      .Source        = "WLC"                      && where did this tool come from?  Your own initials, for instance
      .Sort          = 0                          && the sort order for all items from the same Category
      
      * For public tools or shared tools, such as PEM Editor, etc.
      .Version       = "Version 1.2"                                && e.g., 'Version 7, May 18, 2011'
      .Author        = "Rick Schummer"
      .Link          = "https://github.com/rschummer/ThorTools"     && link to a page for this tool
      .VideoLink     = SPACE(0)                                     && link to a video for this tool
      
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
*  AUTHOR: Richard A. Schummer, May 2021
*
*  METHOD DESCRIPTION:
*    To change the output style to the VFP desktop, in a message box or 
*    wait window as defined by the constant ccISTOTD_MSGSTYLE.
*
*    To change what is detailed for each entry, see the constants ccISTOTD_EXPR1, 
*    ccISTOTD_EXPR2, ccISTOTD_EXPR3, ccISTOTD_EXPR4, and ccISTOTD_EXPR5.
*
*    To change how many Intellisense entries are displayed, see the constant
*    cnISTOTD_HOWMANY
*
*    Additionally, you can localize text displayed in the message via the 
*    constants ccISTOTD_MSGBOX_TITLE_LOC, ccISTOTD_LEADTEXT_LOC, 
*    ccISTOTD_QFC_RECCOUNT_LOC and ccISTOTD_FC_RECCOUNT_LOC
*
*  INPUT PARAMETERS:
*    txParam1 = unknown type, optional, standard parameter passed in by Thor.
*               Not used by this Thor tool.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode(txParam1)

* Captions set up for localization
#DEFINE ccISTOTD_MSGBOX_TITLE_LOC  [Intellisense Tip of the Day]
#DEFINE ccISTOTD_LEADTEXT_LOC      [Intellisense Tip of the Day:]
#DEFINE ccISTOTD_QFC_RECCOUNT_LOC  [xx records pulled from FoxCode for ] + ccISTOTD_MSGBOX_TITLE_LOC
#DEFINE ccISTOTD_FC_RECCOUNT_LOC   [xx records in the FoxCode table]

* Valid values for Message Style: Desktop, MessageBox, or WaitWindow
#DEFINE ccISTOTD_MSGSTYLE          "Desktop"

* Remove Script, Version, COM, Object types, and Special and blank Abbreviations
* and specifically show WLC entries based on docs in the User column 
#DEFINE ccISTOTD_FILTER            [!INLIST(type, "S", "V", "O", "T", "Z") AND !EMPTY(Abbrev) AND "WHITE LIGHT" $ UPPER(user)]

* Define how many entries should be shown, if you want extra white space before first entry display (desktop),
* and the tip information displayed/output.
#DEFINE cnISTOTD_HOWMANY           3
#DEFINE clEXTRA_WHITESPACE         .T.
#DEFINE ccISTOTD_TYPES             ICASE(type="U", "User", type="C", "Command", type="E", "Property Editor", type="M", "MenuHit", type="F", "Function", type="P", "Property", type="S", "Script", type="T", "Object Type", type="O", "COM", type="V", "Version", type="Z", "Special", "Undefined")

#DEFINE ccISTOTD_EXPR1             ALLTRIM(EVALUATE("Abbrev")) + IIF(EMPTY(Expanded), SPACE(0), " -> " + ALLTRIM(EVALUATE("Expanded"))) + IIF(Type = "F", "(" + Tip + ") ", SPACE(0)) + SPACE(1)
#DEFINE ccISTOTD_EXPR2             "[" + ccISTOTD_TYPES + ", " + TRANSFORM(nRecNo) + "]" + ccCRLF 
#DEFINE ccISTOTD_EXPR3             IIF(INLIST(Type, "F", "E") OR EMPTY(Tip), SPACE(0), Tip + ccCRLF)
#DEFINE ccISTOTD_EXPR4             User + IIF(EMPTY(User), SPACE(0), SPACE(1))
#DEFINE ccISTOTD_EXPR5             "(" + TRANSFORM(timestamp) + ")"

* Local decalrations
LOCAL loSession AS Session, ;
      loException AS Exception

LOCAL lnOldDataSessionID, ;
      lcWhere, ;
      lnI, ;
      lcRandomISTOTD, ;
      loException, ;
      lcCode

* Core Intellisense Tip of the Day code
TRY 
   lnOldDataSessionID = SET("Datasession")
   loSession          = CREATEOBJECT("Session")
   
   SET DELETED ON 
   
   SET DATASESSION TO (m.loSession.DataSessionID)
   
   RAND(-1)
   
   USE IN (SELECT("curFC"))
   USE (_foxcode) IN 0 SHARED AGAIN ALIAS curFC
   
   lcWhere = IIF(EMPTY(ccISTOTD_FILTER), SPACE(0), [ WHERE ] + ccISTOTD_FILTER)
   
   * Randomize the list of FoxCode records based on the filtering condition. 
   SELECT RAND() AS nRandom, RECNO() AS nRecNo, curFC.* ;
      FROM curFC ;
      ORDER BY 1 ;
      &lcWhere ;
      INTO CURSOR curRandomFC NOFILTER 

   SELECT curRandomFC
   
   IF LOWER(ccISTOTD_MSGSTYLE) = "desktop" 
      * Make sure if open tool windows, that the desktop output ends up on desktop
      ACTIVATE SCREEN 
      
      * Spacer to separate from other potential messages dumped on screen  
      IF clEXTRA_WHITESPACE
         ? ccCRLF
      ENDIF  
   ENDIF 
   
   * Display record count of the records that match the filter criteria.
   lcQueriedRecordNumberDetails = STRTRAN(ccISTOTD_QFC_RECCOUNT_LOC, "xx", TRANSFORM(RECCOUNT("curRandomFC")))
   lcFoxCodeRecordNumberDetails = STRTRAN(ccISTOTD_FC_RECCOUNT_LOC,  "xx", TRANSFORM(RECCOUNT("curFC")))
   
   IF LOWER(ccISTOTD_MSGSTYLE) = "desktop"
      ? m.lcQueriedRecordNumberDetails 
      ? m.lcFoxCodeRecordNumberDetails
      ? LOWER(_foxcode)
      ? ""
   ENDIF 
      
   * Display as many as are defined.
   IF RECCOUNT("curRandomFC") > 0
      FOR lnI = 1 TO cnISTOTD_HOWMANY
         lcRandomISTOTD = ccISTOTD_LEADTEXT_LOC + ;
                          ccCRLF + ;
                          ccISTOTD_EXPR1 + ;
                          ccISTOTD_EXPR2 + ;
                          ccISTOTD_EXPR3 + ;
                          ccISTOTD_EXPR4 + ;
                          ccISTOTD_EXPR5 + ;
                          ccCRLF 
                         
         *< lcRandomISTOTD = ccISTOTD_LEADTEXT_LOC + ;
                          ccCRLF + ;
                          ccISTOTD_EXPR2 + ;
                          ccISTOTD_EXPR3 + ;
                          ccISTOTD_EXPR4 + ;
                          ccISTOTD_EXPR5 + ;
                          ccCRLF 
         DO CASE
            CASE LOWER(ccISTOTD_MSGSTYLE) = "desktop"
               ? m.lcRandomISTOTD

            CASE LOWER(ccISTOTD_MSGSTYLE) = "messagebox"
               MESSAGEBOX(m.lcRandomISTOTD, ; 
                          0+64, ;
                          ccISTOTD_MSGBOX_TITLE_LOC)
                          
            CASE LOWER(ccISTOTD_MSGSTYLE) = "waitwindow"
               WAIT WINDOW m.lcRandomISTOTD 

            OTHERWISE
               * No other communication mechanism

         ENDCASE

         SKIP IN curRandomFC

         IF EOF("curRandomFC")
            EXIT 
         ENDIF 
      ENDFOR 
   ENDIF              

   USE IN (SELECT("curFC"))
   USE IN (SELECT("curRandomFC"))
   
   SET DATASESSION TO (m.lnOldDataSessionID)
   
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

RETURN 

*: EOF :*  