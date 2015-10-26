******************************************************************************************
*  PROGRAM: Thor_Tool_WLCProjectBuilder.prg
*
*  AUTHOR: Richard A. Schummer, June 2015
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
*     This program runs the WLC Project Builder if available.
*
*  CALLING SYNTAX:
*     DO Thor_Tool_WLCProjectBuilder.prg
*     Thor_Tool_WLCProjectBuilder()
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
*                                 C H A N G E    L O G                                    
*
*    Date     Developer               Version  Description
* ----------  ----------------------  -------  -------------------------------------------
* 06/20/2015  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     2.2      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

LOCAL loTool AS "Tool"

#DEFINE ccCRLF            CHR(13)+CHR(10)
#DEFINE ccPBCLASSLIBRARY  "d:\devvfp8apps\devtools\projecthook\cprojecthook5.vcx"
#DEFINE ccPBCLASS         "frmProjectBuilder"


* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND 'O' = VARTYPE(lxParam1) AND 'thorinfo' == LOWER(lxParam1.Class)
   WITH lxParam1
   
      * Required
      .Prompt          = "WLC Project Builder"           && used in menus
      .StatusBarText   = "Run the WLC Project Builder dialog"  
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2     && a description for the tool
         This program runs the WLC Project Builder dialog.      
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
      .VideoLink     = SPACE(0)                                    && link to a video for this tool
      
   ENDWITH 

   RETURN lxParam1
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

   MESSAGEBOX(lcCode, ;
              0+48, ;
              _screen.Caption)
 
ENDTRY

RETURN


***********************************************************************************
DEFINE CLASS Tool AS Custom

cCaption        = "WLC Project Builder"
cPBClassLibrary = ccPBCLASSLIBRARY
cPBClass        = ccPBCLASS

   ********************************************************************************
   *  METHOD NAME: Do
   *
   *  AUTHOR: Richard A. Schummer,
   *
   *  METHOD DESCRIPTION:
   *    Run the WLC Project Builder, instantiate if necessary.
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
         lcOldSafety

   lcOldSafety = SET("Safety")
   SET SAFETY OFF

   TRY
      IF TYPE("_vfp.ActiveProject") = "O"
         this.InstanceProjectBuilder()
         _screen.__oWLCProjectBuilder.Show()
      ELSE
         MESSAGEBOX("The " + this.cCaption + " requires an open project.", ; 
                    0+48, ;
                    _screen.Caption)
      ENDIF

      
   CATCH TO loException
      lcCode = "Error: " + m.loException.Message + ;
               " [" + TRANSFORM(m.loException.Details) + "] " + ;
               " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
               " in " + m.loException.Procedure + ;
               " on " + TRANSFORM(m.loException.LineNo)

      MESSAGEBOX(lcCode, ;
                 0+48, ;
                 this.cCaption)
    
   ENDTRY

   SET SAFETY &lcOldSafety

   RETURN  
   
   ENDPROC 
   
   
   ********************************************************************************
   *  METHOD NAME: InstanceProjectBuilder
   *
   *  AUTHOR: Richard A. Schummer, August 2015
   *
   *  METHOD DESCRIPTION:
   *    This method is called to create a reference and instance of the WLC Project Builder
   *
   *  INPUT PARAMETERS:
   *    None
   * 
   *  OUTPUT PARAMETERS:
   *    None
   * 
   ********************************************************************************
   PROCEDURE InstanceProjectBuilder()

   IF NOT PEMSTATUS(_screen, "__oWLCProjectBuilder", 5)
      _screen.AddProperty("__oWLCProjectBuilder")
   ENDIF

   IF TYPE("_screen.__oWLCProjectBuilder") = "O" OR ISNULL(_screen.__oWLCProjectBuilder) OR EMPTY(_screen.__oWLCProjectBuilder)
      _screen.__oWLCProjectBuilder = NEWOBJECT(this.cPBClass, this.cPBClassLibrary)
   ENDIF 

   RETURN   
   
   ENDPROC 


   ********************************************************************************
   *  METHOD NAME: Init
   *
   *  AUTHOR: Richard A. Schummer
   *
   *  METHOD DESCRIPTION:
   *    Method to write behavior that Occurs when an object is created.
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
   *    Method to write behavior that Occurs when an object is released.
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

      MESSAGEBOX(lcCode, ;
                 0+48, ;
                 this.cCaption)
      
      RETURN 
   ENDPROC
ENDDEFINE

*: EOF :*