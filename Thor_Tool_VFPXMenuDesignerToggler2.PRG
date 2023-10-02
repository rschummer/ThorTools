******************************************************************************************
*  PROGRAM: Thor_Tool_VfpxMenuDesignerToggler2.prg
*
*  AUTHOR: Richard A. Schummer, August 2023
*
*  COPYRIGHT © 2023   All Rights Reserved.
*     White Light Computing, Inc.
*     PO Box 391
*     Washington Twp., MI  48094
*     raschummer@whitelightcomputing.com
*
*  PROGRAM DESCRIPTION:
*     This Thor tool toggles the _menudesigner system memvar between the VFPX MenuDesigner Pro
*     and the Visual FoxPro Menu Designer. It does this by looking up the location of
*     MenuDesigner Pro in the Windows Registry. If it cannot find it there, it sets the 
*     system memvar to no characters, which effectively sets the Visual FoxPro IDE to 
*     run the Visual FoxPro Menu Designer when editing menus.
*
*     All messaging as to the status of which Menu Designer is in effect is controlled by
*     a property called cMessagingScheme in the Tool class.
*
*  CALLING SYNTAX:
*     Normally called by Thor when it starts up to add this tool to the Thor tool menu.
*     You can also run this program standalone, outside of Thor to accomplish the same
*     thing.
*
*     DO Thor_Tool_VfpxMenuDesignerToggler2.prg
*     Thor_Tool_VfpxMenuDesignerToggler2()
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
*                                C H A N G E    L O G                                     
*
*    Date     Developer               Version  Description
* ----------  ----------------------  -------  -------------------------------------------
* 08/14/2023  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS txParam1

LOCAL loTool AS "Tool"

#DEFINE ccCRLF            CHR(13)+CHR(10)

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND "O" = VARTYPE(txParam1) AND "thorinfo" == LOWER(txParam1.Class)
   WITH txParam1
   
      * Required, used in menus
      .Prompt          = "MenuDesigner Toggler"           
      .StatusBarText   = "Toggle between the VFPX MenuDesigner Pro and VFP MenuDesigner"  
      
      * Optional, a description for the tool
      TEXT TO .Description NOSHOW PRETEXT 1+2
         <<.StatusBarText>>     
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source          = "VFPX"                     && where did this tool come from?  Your own initials, for instance
      .Category        = "VFPX"                     && creates categorization of tools; defaults to .Source if empty
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

* Major version of the MenuDesigner Professional (used for registry keys)
#DEFINE ccMD_VERSION "v3.0"

cCaption         = "MenuDesigner Pro Toggler"
cRegistryKey     = SPACE(0)
cLocation        = SPACE(0)

cMessagingScheme = "W"                  && (M)essageBox, (WW)ait Window, (W)aitWindow NoWait, (D)esktop

oRegistry        = NULL

********************************************************************************
*  METHOD NAME: Do
*
*  AUTHOR: Richard A. Schummer,
*
*  METHOD DESCRIPTION:
*    This method is called to run the tool from Thor menu, hotkey, toolbar button,
*    or favorite Thor launcher.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE Do()

LOCAL loException AS Exception

LOCAL lcMajorVersion, ;
      lcCode

TRY
   IF EMPTY(_menudesigner) 
      lcMajorVersion = SUBSTRC(ccMD_VERSION, 1, ATC(".", ccMD_VERSION)) + "0"
      lcMajorVersion = STRTRAN(lcMajorVersion, "v", SPACE(0))
      lcMajorVersion = STRTRAN(lcMajorVersion, "V", SPACE(0))

      * Build the registry key to show in the ToolTip text
      this.cRegistryKey = "Software\WhiteLightComputingTools\MenuDesigner\" + ;
                          lcMajorVersion + ;
                          "\Options"
   
      this.GetLocationRegistryEntry()
      
      IF EMPTY(this.cLocation) OR ISNULL(this.cLocation)
         * Nothing found
         this.CustomMessageBox("WARNING: The registry entry for MenuDesigner Pro was not found or is not set up. " + ;
                               "Best solution is to register it using MenuDesigner Pro via the Configuration tab.", ; 
                               0+48, ;
                               this.cCaption)
      ELSE
         IF FILE(this.cLocation)
            _menudesigner = ALLTRIM(this.cLocation)
            
            this.CustomMessageBox("MenuDesigner Pro is now set as the menu designer (" + _menudesigner + ").", ; 
                                  0+64, ;
                                  this.cCaption)
         ELSE
            this.CustomMessageBox("WARNING: MenuDesigner Pro registry has incorrect location, please re-register location using Configuration tab of the MenuDesigner Pro.", ; 
                                  0+64, ;
                                  this.cCaption)
         ENDIF 
      ENDIF 
   ELSE
      _menudesigner = SPACE(0)

      this.CustomMessageBox("Native VFP Menu Designer is now set as the menu designer.", ; 
                            0+64, ;
                            this.cCaption)
   ENDIF 

   
CATCH TO loException
   lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   this.CustomMessageBox(m.lcCode, ;
              0+48, ;
              this.cCaption)
 
ENDTRY

RETURN  

ENDPROC 


********************************************************************************
*  METHOD NAME: CustomMessageBox
*
*  AUTHOR: Richard A. Schummer, August 2023
*
*  METHOD DESCRIPTION:
*    Display messages to the developer. You determine the scheme of messaging used.
*       1) (M)essageBox
*       2) (W)ait (W)indow 
*       2) (W)ait Window (no wait)
*       3) (D)esktop screen
*
*    Make your selection by setting the cMessagingScheme property for this class.
*
*  INPUT PARAMETERS:
*    tcMessageText   = Required, character, text for the message being displayed.
*    tnDialogBoxType = For MessageBox, numeric, same as second parameter
*                      for MESSAGEBOX(), defaults to zero if nothing passed, which 
*                      is just the OK button with no icons. Specifies the buttons 
*                      and icons that appear in the dialog box, the default button 
*                      when the dialog box is displayed, and the behavior of the 
*                      dialog box.
*    tcTitleBarText  = For MessageBox, character, defaults to _screen.Caption if 
*                      nothing is passed. Specifies the text that appears in the 
*                      title bar of the dialog box. 
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CustomMessageBox(tcMessageText, tnDialogBoxType, tcTitleBarText)

LOCAL lcMessagingScheme

lcMessagingScheme = this.cMessagingScheme

IF VARTYPE(m.tcMessageText) # "C"
   m.tcMessageText = "Message parameter passed to display message is invalid, please debug"
ENDIF 

IF m.lcMessagingScheme = "M"
   IF VARTYPE(m.tnDialogBoxType) # "N"
       m.tnDialogBoxType = 0                 && Plain OK button, no icons
   ENDIF 
   
   IF VARTYPE(m.tcTitleBarText) # "C"
       m.tcTitleBarText = _screen.Caption    && Default to the current VFP caption.
   ENDIF 
ENDIF 

DO CASE
   CASE m.lcMessagingScheme = "M"
      MESSAGEBOX(m.tcMessageText, ; 
                 m.tnDialogBoxType, ;
                 m.tcTitleBarText)

   CASE m.lcMessagingScheme = "WW"
      WAIT WINDOW (m.tcMessageText)

   CASE m.lcMessagingScheme = "W"
      WAIT WINDOW (m.tcMessageText) NOWAIT

   CASE m.lcMessagingScheme = "D"
      ACTIVATE SCREEN 
      ? m.tcMessageText

   OTHERWISE
      * Nothing to do, valid settings are enforced in the property access method
      * and defaults to Wait Windows if invalid setting is attempted.

ENDCASE

RETURN 

ENDPROC 


********************************************************************************
*  METHOD NAME: cMessagingScheme_Access
*
*  AUTHOR: Richard A. Schummer, August 2023
*
*  METHOD DESCRIPTION:
*    This is the access method code for the cMessagingScheme property used to 
*    identify how the messaging is done for this tool.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE cMessagingScheme_Access()

IF INLIST(this.cMessagingScheme, "M", "WW", "W", "D")
   * All set, valid setting
ELSE
   this.cMessagingScheme = "W"
ENDIF 

RETURN UPPER(this.cMessagingScheme)


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

RETURN 

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
this.oRegistry = NULL

RETURN 

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

LOCAL lcCode

DIMENSION laError[1]

AERROR(laError)

lcCode = "Error: " + laError[2] + ;
         " (" + TRANSFORM(laError[1]) + ")" + ;
         " in " + m.tcMethod + ;
         " on " + TRANSFORM(m.tnLine)

this.CustomMessageBox(m.lcCode, ;
                      0+48, ;
                      this.cCaption)

RETURN 

ENDPROC


********************************************************************************
*  METHOD NAME: GetLocationRegistryEntry
*
*  AUTHOR: Richard A. Schummer, August 2023
*
*  METHOD DESCRIPTION:
*    This method is called to get the registry entry for the MenuDesigner location.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE GetLocationRegistryEntry()

* Registry roots (ripped off from VFP\ffc\Registry.h)
#DEFINE HKEY_CLASSES_ROOT           -2147483648  && BITSET(0,31)
#DEFINE HKEY_CURRENT_USER           -2147483647  && BITSET(0,31)+1
#DEFINE HKEY_LOCAL_MACHINE          -2147483646  && BITSET(0,31)+2
#DEFINE HKEY_USERS                  -2147483645  && BITSET(0,31)+3

LOCAL lcToolKey, ;
      lcLocation

this.GetRegistryInstance()

* Only process the registry entries if Registry object instantiated
IF NOT ISNULL(this.oRegistry)
   lcToolKey  = this.cRegistryKey
   lcLocation = SPACE(0)

   this.oRegistry.GetRegKey("Location", ;
                            @lcLocation,;
                            m.lcToolKey,;
                            HKEY_CURRENT_USER)

   this.cLocation = LOWER(lcLocation)
ENDIF

RETURN 

ENDPROC


********************************************************************************
*  METHOD NAME: GetRegistryInstance
*
*  AUTHOR: Richard A. Schummer, August 2023
*
*  METHOD DESCRIPTION:
*    This method is called to instantiate the registry object from the registry 
*    class library in the Fox Foundation Class (FFC).
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE GetRegistryInstance()

LOCAL lcRegistryClassLib

IF ISNULL(this.oRegistry)
   lcRegistryClassLib = HOME()+"FFC\registry.vcx"

   this.oRegistry     = NEWOBJECT("registry", m.lcRegistryClassLib)
ENDIF

RETURN this.oRegistry

ENDPROC

ENDDEFINE

*: EOF :*      