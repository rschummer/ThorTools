******************************************************************************************
*  PROGRAM: Thor_Tool_CreateThorToolWOThorUI.prg
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
*     This program creates a Thor tool from selected template without the Thor Configuration dialog.
*
*  CALLING SYNTAX:
*     DO Thor_Tool_CreateThorToolWOThorUI.prg
*     Thor_Tool_CreateThorToolWOThorUI()
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
* 06/05/2015  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     1.1      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

#DEFINE ccCRLF            CHR(13)+CHR(10)

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND 'O' = VARTYPE(lxParam1) AND 'thorinfo' == LOWER(lxParam1.Class)
   WITH lxParam1
   
      * Required
      .Prompt          = "Create Thor Tool without Thor UI"       && used in menus
      .StatusBarText   = .Prompt  
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2                         && a description for the tool
         This program creates a Thor Tool without using the Thor Configuration user interface.      
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

IF PCOUNT() = 0
   DO ToolCode
ELSE
   DO ToolCode WITH lxParam1
ENDIF

RETURN


********************************************************************************
*  METHOD NAME: ToolCode
*
*  AUTHOR: Richard A. Schummer,
*
*  METHOD DESCRIPTION:
*    Main tool processing code.
*
*  INPUT PARAMETERS:
*    lxParam1 = unknown type, optional, standard parameter passed in by Thor.
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ToolCode
LPARAMETERS lxParam1

#DEFINE ccCAPTION   "WLC Create Thor Tool"

LOCAL loException as Exception, ;
      lcOldSafety, ;
      loCboTemplates AS ComboBox

lcOldSafety = SET("Safety")
SET SAFETY OFF

TRY
   lcToolsFolder      = ALLTRIM(_screen.oThorEngine.cToolsFolder)
   lcProcsFolder      = ADDBS(ADDBS(lcToolsFolder) + "Procs")
   lcMyToolsFolder    = ADDBS(ADDBS(lcToolsFolder) + "My Tools")
   lcMyTemplateFolder = ADDBS(ADDBS(lcMyToolsFolder) + "Templates")
   lcBaseProgFileName = "Thor_Tool_"
   lcNewProgFileName  = lcBaseProgFileName
   
   lnTemplates        = ADIR(laTemplate, lcMyTemplateFolder + "*.prg")
   lnSorted           = ASORT(laTemplate, -1, -1, 0, 1)
   
   IF lnTemplates > 0
      * Have template default to first in the list, unless there is one named as with DEFAULT at the end of the file name stem.
      lcTemplateSelected = laTemplate[1,1]

      FOR lnI = 1 TO lnTemplates 
         IF RIGHTC(UPPER(JUSTSTEM(laTemplate[lnI, 1])), 7) = "DEFAULT"
            lcTemplateSelected = laTemplate[lnI, 1]
            EXIT 
         ENDIF 
      ENDFOR 

      loForm = NEWOBJECT("DynamicForm", ADDBS(lcProcsFolder) + "Thor_Proc_DynamicForm.PRG")

      loForm.Caption                    = ccCaption
      loForm.cHeading                   = "Please note this tool is not available until you open and close Thor Configuration, or click on Thor button to re-register tools."
      loForm.nHeadingFontSize           = 10
      loForm.oRenderEngine.lLabelsAbove = .T.
      
      TEXT TO lcBodyMarkup NOSHOW TEXTMERGE
         lcTemplateSelected :Enabled => .T.
                            :Class => "ComboBox"
                            :ColumnCount => 4
                            :BoundColumn => 1
                            :ColumnWidths => "500, 100, 100"
                            :RowSource => "laTemplate"
                            :RowSourceType => 5 
                            :Width => 500
                            :Label.FontBold => .T. 
                            :Label.Caption => "Templates" |

         lcNewProgFileName  :Class => "TextBox"
                            :Width => 400 
                            :Label.FontBold => .T. 
                            :Label.Caption => "Tool Name" |
      ENDTEXT
      
      * Set the main body area markup 
      loForm.cBodyMarkup = lcBodyMarkup

      * Call Render method to create control in the Form
      llRenderedOK = loForm.Render()

      * Show the form to the user
      IF llRenderedOK
         * Note. You have a chance here to programmatically change anything on the form or controls
         * in any way needed before showing the form to the user...
         
         loForm.AutoCenter = .T.
         loForm.Show(1)               
      ELSE 
         MESSAGEBOX(loForm.oRenderEngine.GetErrorsAsString() , 0, 'Notice.')
         * If there were any rendering errors (llResult = .f.), then you can read loForm.oRenderEngine.nErrorCount property
         * and loForm.oRenderEngine.oErrors collection for a detail of each error. Or call loForm.oRenderEngine.GetErrorsAsString().
         loForm.Show(1)
      ENDIF 

      * At this point, the user is interacting with the form, and it will eventually be closed when they click
      * Save, Cancel, or the [X] button. At that time, flow will return here, and we can then ready any property
      * on loForm and  loForm.oRenderEngine, and even access the rendered controls.
      IF VARTYPE(loForm) = 'O' AND LOWER(loForm.cReturn) = 'save'
         *-- If Save is clicked, the controlsources are already updated with the new values from the UI.
         *-- Do whatever local processing you need following the Save click by the user...
         
         lcNewProgFileName = FORCEEXT(ALLTRIM(lcNewProgFileName), "prg") 
         
         IF lcNewProgFileName == FORCEEXT(lcBaseProgFileName, "prg")
            * Nothing to create as developer did not change the name.
         ELSE
            IF FILE(lcMyToolsFolder + lcNewProgFileName)
               MESSAGEBOX("The file name you want to create already exists. Please try again.", ; 
                          0+16, ;
                          _screen.Caption)
            ELSE
               COPY FILE (lcMyTemplateFolder + lcTemplateSelected) TO ;
                         (lcMyToolsFolder + lcNewProgFileName)
                         
               MODIFY COMMAND (lcMyToolsFolder + lcNewProgFileName) NOWAIT 
            ENDIF 
         ENDIF  
      ELSE 
         *-- Do whatever processing for Close/Cancel user action...
         *-- If using the Button Bar or and instance of DF_CancelButton on the form and Cancel was clicked,
         *-- and the property loForm.lRestoreDataOnCancel = .t. (default), then the controlsources will already
         *-- be restored to their original value by the Form class.
      ENDIF 
      
      
      * After the preceding Save/Cancel processing, we can now Release the loForm object.
      RELEASE loForm

   ELSE
      MESSAGEBOX("No templates found in " + lcMyTemplateFolder + ".", ; 
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
              _screen.Caption)
 
ENDTRY

SET SAFETY &lcOldSafety

RETURN  

*: EOF :* 