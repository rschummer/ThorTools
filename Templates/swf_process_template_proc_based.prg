LPARAMETERS txParam1

#DEFINE ccCRLF       CHR(13)+CHR(10)

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND "O" = VARTYPE(txParam1) AND "thorinfo" == LOWER(txParam1.Class)
   WITH txParam1
   
      * Required and used in menus
      .Prompt          = "<FILL ME IN>"           
      .StatusBarText   = "This process for Geek Gathering conferences "  
      
      * Optional, a description for the tool, easy to overwrite with addtional text, even after the status bar text.
      TEXT TO .Description TEXTMERGE NOSHOW PRETEXT 1+2 
         <<.StatusBarText>>      
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Category        = "Southwest Fox"                           && creates categorization of tools; defaults to .Source if empty
      .Source          = "WLC"                                     && where did this tool come from?  Your own initials, for instance
      .Sort            = 0                                         && the sort order for all items from the same Category
      
      * For public tools or shared tools, such as PEM Editor, etc.
      .Version         = "Version 1.0"                             && e.g., 'Version 7, May 18, 2011'
      .Author          = "Rick Schummer"
      .Link            = "https://github.com/rschummer/ThorTools"  && link to a page for this tool
      .VideoLink       = SPACE(0)                                  && link to a video for this tool
      
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
*  AUTHOR: Richard A. Schummer
*
*  METHOD DESCRIPTION:
*    
*
*  INPUT PARAMETERS:
*    txParam1 = unknown type, optional, standard parameter passed in by Thor.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode(txParam1)

LOCAL loException as Exception, ;
      lcOldSafety

lcOldSafety = SET("Safety")
SET SAFETY OFF

TRY
   CD j:\wlcproject\southwestfox\eventmanagementsystem
   DO progs\
   
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

RETURN  

*: EOF :* 