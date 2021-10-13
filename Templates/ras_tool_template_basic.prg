LPARAMETERS lxParam1

#DEFINE ccCRLF            CHR(13)+CHR(10)

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND "O" = VARTYPE(lxParam1) AND "thorinfo" == LOWER(lxParam1.Class)
   WITH lxParam1
   
      * Required, used in menus
      .Prompt          = "<FILL ME IN>" 
      .StatusBarText   = "<FILL ME IN>"  
      
      * Optional, a description for the tool
      TEXT TO .Description NOSHOW PRETEXT 1+2 
         <<.StatusBarText>>      
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Category        = "WLC"                               && creates categorization of tools; defaults to .Source if empty
      .Source          = "WLC"                               && where did this tool come from?  Your own initials, for instance
      .Sort            = 0                                   && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version         = "Version 1.0"                                  && e.g., 'Version 7, May 18, 2011'
      .Author          = "Rick Schummer"
      .Link            = "https://github.com/rschummer/ThorTools"       && link to a page for this tool
      .VideoLink       = SPACE(0)                                       && link to a video for this tool
      
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
*  AUTHOR: Richard A. Schummer, November 2014
*
*  METHOD DESCRIPTION:
*    
*    
*
*  INPUT PARAMETERS:
*    lxParam1 = unknown type, standard parameter passed in by Thor.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode


 
RETURN  

*: EOF :* 
