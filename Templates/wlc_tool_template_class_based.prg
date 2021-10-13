LPARAMETERS txParam1

LOCAL loTool AS "Tool"

#DEFINE ccCRLF            CHR(13)+CHR(10)

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND "O" = VARTYPE(txParam1) AND "thorinfo" == LOWER(txParam1.Class)
   WITH txParam1
   
      * Required, used in menus
      .Prompt          = "<FILL ME IN>"
      .StatusBarText   = "<FILL ME IN>"  
      
      * Optional, a description for the tool
      TEXT TO .Description NOSHOW PRETEXT 1+2
         <<.StatusBarText>>     
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source          = "WLC"                      && where did this tool come from?  Your own initials, for instance
      .Category        = "WLC"                      && creates categorization of tools; defaults to .Source if empty
      .Sort            = 0                          && the sort order for all items from the same Category
      
      * For public or shared tools, such as PEM Editor, etc.
      .Version         = "Version 1.0"                             && e.g., 'Version 7, May 18, 2011'
      .Author          = "Rick Schummer"
      .Link            = "https://github.com/rschummer/ThorTools"  && link to a page for this tool
      .VideoLink       = SPACE(0)                                  && link to a video for this tool
      
   ENDWITH 

   RETURN txParam1
ENDIF 

TRY 
   loTool = CREATEOBJECT("Tool")
   loTool.Do()
   loTool.Release()

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


***********************************************************************************
DEFINE CLASS Tool AS Custom

cCaption = "Thor-based tool"

********************************************************************************
*  METHOD NAME: Do
*
*  AUTHOR: Richard A. Schummer,
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
PROCEDURE Do()

LOCAL loException as Exception, ;
      lcOldSafety, ;
      lcCode

lcOldSafety = SET("Safety")
SET SAFETY OFF

TRY

   
CATCH TO loException
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(m.lcCode, ;
              0+48, ;
              this.cCaption)
 
ENDTRY

SET SAFETY &lcOldSafety

RETURN  

ENDPROC 


********************************************************************************
*  METHOD NAME: Init
*
*  AUTHOR: Richard A. Schummer
*
*  METHOD DESCRIPTION:
*    Method to write behavior that occurs when an object is created and initialized.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE Init()

ENDPROC


********************************************************************************
*  METHOD NAME: Destroy
*
*  AUTHOR: Richard A. Schummer
*
*  METHOD DESCRIPTION:
*    Method to write behavior that occurs when an object is released.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE Destroy()
* Release all object properties created by the class.

ENDPROC


********************************************************************************
*  METHOD NAME: Release
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    Generic method to call to release the object.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE Release()

RELEASE this

RETURN 

ENDPROC 


********************************************************************************
*  METHOD NAME: Error
*
*  AUTHOR: Richard A. Schummer, June 2015
*
*  METHOD DESCRIPTION:
*    Standard error method for the class. Displays a messagebox for the developer
*    to see what error occured, and what line of what method. Defined here in case
*    the class has an error in it before the Do method is executed. The Do method 
*    has wrapper of TRY...END so developer has more control over the error handling.
*    Not all that useful if the class is instantiated inside a TRY...END wrap.
*
*  INPUT PARAMETERS:
*    tnError  = numeric, the number of the error. Identical to the value returned by ERROR( ).
*    tcMethod = character, the method name where the error occured.
*    tnLine   = numeroc, the line number of the method where the error occured.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE Error(tnError, tcMethod, tnLine)

AERROR(laError)

lcCode = "Error: " + laError[2] + ;
         " (" + TRANSFORM(laError[1]) + ")" + ;
         " in " + m.tcMethod + ;
         " on " + TRANSFORM(m.tnLine)

MESSAGEBOX(m.lcCode, ;
           0+48, ;
           this.cCaption)

RETURN 

ENDPROC

ENDDEFINE

*: EOF :*