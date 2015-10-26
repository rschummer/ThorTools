******************************************************************************************
*  PROGRAM: Thor_Tool_ProjectStatistics.prg
*
*  AUTHOR: Richard A. Schummer, July 2015
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
*     This program produces statistics about the selected project.
*
*  CALLING SYNTAX:
*     DO Thor_Tool_ProjectStatistics.prg
*     Thor_Tool_ProjectStatistics()
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
* 07/04/2015  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     1.1      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS txParam1

#DEFINE ccCRLF                CHR(13)+CHR(10)
#DEFINE ccTAB                 CHR(9)

* Default settings for the tool options.
#DEFINE ccTOOLNAME            "WLC Project Statistics"
#DEFINE clCOPYTOFILE          .T.
#DEFINE ccPROJECTSTATLOGPATH  ADDBS(SYS(2023))
#DEFINE ccPROJECTSTATLOGFILE  [WLCProjectStatsLog.txt]

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND 'O' = VARTYPE(m.txParam1) AND 'thorinfo' == LOWER(m.txParam1.Class)
   WITH m.txParam1
   
      * Required and used in menus
      .Prompt          = ccTOOLNAME           
      .StatusBarText   = "Analyze and produce statistics about the active project in the IDE"  
      
      * Optional, a description for the tool, easy to overwrite with addtional text, even after the status bar text.
      TEXT TO .Description TEXTMERGE NOSHOW PRETEXT 1+2 
         <<.StatusBarText>>      
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * Options Dialog settings
      .OptionTool      = ccTOOLNAME
      .OptionClasses   = "cusShowStatsFile, cusStatsFileName"

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Category      = "WLC"                                       && creates categorization of tools; defaults to .Source if empty
      .Source        = "WLC"                                       && where did this tool come from?  Your own initials, for instance
      .Sort          = 0                                           && the sort order for all items from the same Category
      
      * For public tools or shared tools, such as PEM Editor, etc.
      .Version       = "Version 1.1, September 20, 2015"           && e.g., 'Version 7, May 18, 2011'
      .Author        = "Rick Schummer"
      .Link          = "https://github.com/rschummer/ThorTools"    && link to a page for this tool
      .VideoLink     = SPACE(0)                                    && link to a video for this tool
      
   ENDWITH 

   RETURN m.txParam1
ENDIF 

IF PCOUNT() = 0
   DO ToolCode
ELSE
   DO ToolCode WITH m.txParam1
ENDIF

RETURN


********************************************************************************
*  METHOD NAME: ToolCode
*
*  AUTHOR: Richard A. Schummer,
*
*  METHOD DESCRIPTION:
*    This method drives the tool generation code.
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
      lcOldSafety, ;
      lcCode, ;
      llToFile, ;
      lnFiles

m.lcOldSafety = SET("Safety")
SET SAFETY OFF

TRY
   IF TYPE("_vfp.ActiveProject") = "O"
      CloseCursors()
      CreateFileTypeCursor()
      CreateProjectStatsCursor()
   
      USE (_vfp.ActiveProject.Name) IN 0 SHARED AGAIN NOUPDATE ALIAS _pjxanalysis
      
      * Folder Counts
      SELECT DISTINCT ;
             LOWER(PADR(JUSTPATH(FULLPATH(name)),250)) AS cName, ;
             COUNT(*) AS Count ;
         FROM _pjxanalysis ; 
         GROUP BY cName ;
         ORDER BY cName ;
         INTO CURSOR curCountByFolders READWRITE
         
      INSERT INTO curCountByFolders (cName, Count) VALUES ("TOTAL Folder Count", RECCOUNT("curCountByFolders"))
      GO TOP IN curCountByFolders
   
      * File Counts
      SELECT DISTINCT ;
             curFileType.Description, ;
             curFileType.SortOrder, ;
             COUNT(*) as Count ;
         FROM _pjxanalysis ; 
            JOIN curFileType ON _pjxanalysis.Type = curFileType.Type ;
         WHERE curFileType.Description # "Header" ;
           AND curFileType.Description # "Icon" ;
         GROUP BY curFileType.Description, curFileType.SortOrder ;
         ORDER BY curFileType.SortOrder ;
         INTO CURSOR curCountByFileType READWRITE
         
      SUM curCountByFileType.Count TO m.lnFiles
      
      INSERT INTO curCountByFileType (Description, Count) VALUES ("TOTAL File Count", m.lnFiles)
      
      GO TOP IN curCountByFileType

      llToFile  = EXECSCRIPT(_Screen.cThorDispatcher, "Get Option=", "Copy Stats to File", ccTOOLNAME)

      IF ISNULL(m.llToFile) OR NOT m.llToFile
         BrowseStats()
      ELSE
         TextFileStats()
      ENDIF      
      
      USE IN (SELECT("_pjxanalysis"))
   ELSE
      MESSAGEBOX("The " + ccTOOLNAME + " requires an open project.", ; 
                 0+48, ;
                 _screen.Caption)
   ENDIF 
   
CATCH TO m.loException
   m.lcCode = "Error: " + m.loException.Message + ;
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


********************************************************************************
*  METHOD NAME: BrowseStats
*
*  AUTHOR: Richard A. Schummer, July 2015
*
*  METHOD DESCRIPTION:
*    Display statistics collected in a set of BROWSE windows.
*    
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE BrowseStats()

* Display for the developer
SELECT curCountByFolders
BROWSE LAST NOCAPTIONS NOWAIT TITLE "Project File Counts By Folder"

SELECT curCountByFileType
BROWSE FIELDS Description, Count LAST NOCAPTIONS NOWAIT TITLE "Project File Counts By Type"

SELECT curProjectStats
BROWSE FIELDS Description, Value LAST NOCAPTIONS NOWAIT TITLE "Project Settings"

RETURN 


********************************************************************************
*  METHOD NAME: BrowseStats
*
*  AUTHOR: Richard A. Schummer, July 2015
*
*  METHOD DESCRIPTION:
*    Display statistics collected in a text editor.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE TextFileStats()

LOCAL lcText

m.lcText = SPACE(0)

lcText = m.lcText + ccTOOLNAME + ccCRLF
lcText = m.lcText + TRANSFORM(DATETIME()) + ccCRLF + ccCRLF


* Display for the developer
SELECT curProjectStats

CALCULATE MAX(LENC(ALLTRIM(Description))) TO lnMaxLen IN curProjectStats

lcText = m.lcText + PADR("Project Settings ", 100, "-") + ccCRLF + ccCRLF

SCAN 
   lcText = m.lcText + ;
            PADR(curProjectStats.Description, lnMaxLen + 3) +  ;
            ALLTRIM(curProjectStats.Value) + ccCRLF
ENDSCAN

SELECT curCountByFileType

CALCULATE MAX(LENC(ALLTRIM(Description))) TO lnMaxLen IN curCountByFileType

lcText = m.lcText + ccCRLF
lcText = m.lcText + PADR("Project File Counts By Type ", 100, "-") + ccCRLF + ccCRLF

SCAN 
   lcText = lcText + ;
            PADR(curCountByFileType.Description, lnMaxLen + 3) + ;
            TRANSFORM(curCountByFileType.Count) + ccCRLF
ENDSCAN

 
SELECT curCountByFolders

CALCULATE MAX(LENC(ALLTRIM(cName))) TO lnMaxLen IN curCountByFolders

lcText = m.lcText + ccCRLF
lcText = m.lcText + PADR("Project File Counts By Folder ", 100, "-") + ccCRLF + ccCRLF
  
SCAN 
   lcText = m.lcText + ;
            PADR(curCountByFolders.cName, lnMaxLen + 3) + ;
            TRANSFORM(curCountByFolders.Count) + ccCRLF
ENDSCAN

lcText = m.lcText + ccCRLF
lcText = m.lcText + "*: EOF :*" + ccCRLF


lcFileName = EXECSCRIPT(_Screen.cThorDispatcher, "Get Option=", "Log File Name", ccTOOLNAME)
lcFileName = IIF(ISNULL(m.lcFileName) OR EMPTY(m.lcFileName), ADDBS(ccPROJECTSTATLOGPATH)+ccPROJECTSTATLOGFILE, m.lcFileName)


STRTOFILE(m.lcText, m.lcFileName, .F.)
MODIFY FILE (m.lcFileName) RANGE 1,1 NOWAIT 
      
RETURN


********************************************************************************
*  METHOD NAME: CloseCursors
*
*  AUTHOR: Richard A. Schummer, August 2015
*
*  METHOD DESCRIPTION:
*    Close all the cursors opened by a previous run of this tool.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CloseCursors()

USE IN (SELECT("curProjectStats"))
USE IN (SELECT("curCountByFileType"))
USE IN (SELECT("curCountByFolders"))
USE IN (SELECT("_pjxanalysis"))

RETURN 


********************************************************************************
*  METHOD NAME: CreateProjectStatsCursor
*
*  AUTHOR: Richard A. Schummer, August 2015
*
*  METHOD DESCRIPTION:
*    This method is called to create and fill the curProjectStats. Record added for
*    each of the VFP Project properties supported in the project object for 
*    Visual FoxPro. Details include the a description, the corresponding value, 
*    and a column to sort to the order you prefer.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CreateProjectStatsCursor()

CREATE CURSOR curProjectStats ;
   (Description  C(100), ;
    Value        c(200), ;
    SortOrder    i ;
   )

WITH _vfp.ActiveProject 
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Project", .Name, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Home Directory", .HomeDir, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version", .VersionNumber, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Built", TRANSFORM(.BuildDateTime), 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("File Count", TRANSFORM(.Files.Count), 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("AutoIncrement?", IIF(.AutoIncrement, "True", "False"), 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Debug?", IIF(.Debug, "True", "False"), 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Encrypted?", IIF(.Encrypted, "True", "False"), 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Main File", .MainFile, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Main Class", .MainClass, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version - Comments", .VersionComments, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version - Company", .VersionCompany, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version - Copyright", .VersionCopyright, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version - Description", .VersionDescription, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version - Product", .VersionProduct, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version - Trademarks", .VersionTrademarks, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Version - Language", .VersionLanguage, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("ProjectHook Class", .ProjectHookClass, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("ProjectHook Library", .ProjectHookLibrary, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("SCC Provider", .SCCProvider, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Icon", .Icon, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Server Project", .ServerProject, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Server Help File", .ServerHelpFile, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Type Lib Name", .TypeLibName, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Type Lib Description", .TypeLibDesc, 000)
   INSERT INTO curProjectStats (Description, Value, SortOrder) VALUES ("Type Lib CLSID", .TypeLibCLSID, 000)
ENDWITH 

GO TOP IN curProjectStats

RETURN 


********************************************************************************
*  METHOD NAME: CreateFileTypeCursor
*
*  AUTHOR: Richard A. Schummer, August 2015
*
*  METHOD DESCRIPTION:
*    This method is called to create and fill the curFileType. Record added for
*    each of the file types supported in the project file for Visual FoxPro.
*    Details include the Type column from the project file, the natural language
*    description and a column to sort to the order you prefer.
*
*  INPUT PARAMETERS:
*    None
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE CreateFileTypeCursor()

CREATE CURSOR curFileType ;
   (Type        C(1), ;
    Description C(100), ;
    SortOrder   i ;
   )
    
* Details found in the Visual FoxPro FileSpec project: _60Pjx.DBF
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("H", "Header", 05)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("K", "Form", 20)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("V", "Visual Class Library", 30)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("P", "Program", 10)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("M", "Menu", 40)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("Q", "Query", 90)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("R", "Report", 50)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("B", "Label", 55)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("L", "Library", 60)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("F", "Format", 95)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("d", "Database", 70)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("D", "Free Table", 75)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("t", "Associated Table", 78)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("I", "Index", 79)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("x", "Other File", 80)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("T", "Text File", 85)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("Z", "Application", 65)
INSERT INTO curFileType (Type, Description, SortOrder) VALUES ("i", "Icon", 87)

RETURN 


***********************************************************************************************************************
* Option classes: (notes from Mike Potjer and Jim Nelson)
*   - Tool property: corresponds to the OptionTool property.
*   - Key property (combined with the Tool property): is what you will use to retrieve a specific setting from the
*     settings table.
*   - Value property: is the initial value for the setting. This does not have to be a string, but can be a data type
*     appropriate for the setting.
*   - EditClassName property: specifies the visual class displayed to the user for editing the setting.
*      + The edit class may, and often does, contain controls for more than one setting.
*      + The edit class does NOT have to be defined in the same .PRG as the option class. To use an edit class
*        defined in some other file, use this syntax:
*            EditClassName = ' clsMyToolOptionsPage of MySharedToolOptions.PRG'
*
*   - An option class MUST be defined in the same PRG as the tool that specifies it in its OptionClasses property. In
*     general, it does not work to subclass an option class from another file, because the class libraries used by Thor
*     are not typically in the path.
*   - Multiple tools may use the same setting (not just the tool that defined it), but they must all specify the same
*     Tool and Key combination when reading from or writing to that setting.
*   - It is not necessary to define an option class in every tool that needs access to a particular setting. However, at
*     least one option class must be defined for a tool if you want an “Options” link to appear on the tool’s page in the
*     Tool Launcher.
***********************************************************************************************************************
DEFINE CLASS cusProjectStatsBase AS Custom
   Tool          = ccTOOLNAME
   EditClassName = "ctrProjectStatsOptionsPage" 
ENDDEFINE

DEFINE CLASS cusShowStatsFile AS cusProjectStatsBase
   Key   = "Copy Stats to File"
   Value = clCOPYTOFILE
ENDDEFINE

DEFINE CLASS cusStatsFileName AS cusProjectStatsBase
   Key   = "Log File Name"
   Value = ccPROJECTSTATLOGPATH + ccPROJECTSTATLOGFILE
ENDDEFINE


*********************************************************************
*-- Define the option page for all the SetPath options.
*********************************************************************
DEFINE CLASS ctrProjectStatsOptionsPage AS Container

PROCEDURE Init()
   LOCAL loRenderEngine

   * Get a reference to the Thor OptionRenderEngine (Dynamic Forms)
   m.loRenderEngine = EXECSCRIPT( _Screen.cThorDispatcher, "Class= OptionRenderEngine" )

   
   * Create the Dynamic Forms Options dialog markup syntax
   TEXT TO m.loRenderEngine.cBodyMarkup NOSHOW TEXTMERGE
      .Class         = "CheckBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .BackStyle     = 0
      .Caption       = "Copy statistics to a file (checked) or BROWSE (unchecked)?"
      .Width         = 350
      .cTool         = ccToolName
      .cKey          = "Copy Stats to File"
      |

      .Class         = "lblLogFile"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Caption       = "Open Last Project Stats File: "
      .WordWrap      = .F.
      .Width         = 250
      |

      .Class         = "TextBox"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .Name          = "txtLogPath"
      .Width         = 300
      .cTool         = ccToolName
      .cKey          = "Log File Name"
      |

      .Class         = "cmdLogFilePicker"
      .FontName      = "Segoe UI"
      .FontSize      = 9
      .row-increment = '0'     
      |

   ENDTEXT

   * Render this container.
   m.loRenderEngine.Render(This, ccToolName )
ENDPROC

ENDDEFINE


*************************************************************************************
DEFINE CLASS cmdLogFilePicker AS CommandButton

Caption  = "..."
Width    = 24

PROCEDURE Click()

LOCAL lcOldDir, ;
      lcFileName

TRY 
   m.lcOldDir   = FULLPATH(CURDIR())
   
   CD (JUSTPATH(this.Parent.txtLogPath.Value))
   
   m.lcFileName = GETFILE("txt", "Log File")
   
   CD (m.lcOldDir)
   
   IF EMPTY(m.lcFileName)
      * Nothing to do
   ELSE
      m.lcLogFile = LOWER(ALLTRIM(m.lcFileName))
      
      * Set to object value for physical presentation, then save to the registry to be permanent.
      this.Parent.txtLogPath.Value = m.lcLogFile

      ExecScript(_Screen.cThorDispatcher, "Set Option=", "Log File Name", ccTOOLNAME, m.lcLogFile)
   ENDIF 
   
CATCH TO m.loException
   m.lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(m.lcCode, ;
              0+48, ;
              _screen.Caption)
 
ENDTRY

ENDPROC

ENDDEFINE


*************************************************************************************
DEFINE CLASS lblLogFile AS Label

PROCEDURE Init()

this.ForeColor     = RGB(0,0,255)
this.MousePointer  = 15
this.FontUnderline = .T.
this.AutoSize      = .T.
   
ENDPROC 

PROCEDURE Click()
TRY 
   MODIFY FILE (this.Parent.txtLogPath.Value) NOEDIT RANGE 0,0 

CATCH TO m.loException   
   m.lcCode = "Error: " + m.loException.Message + ;
            " [" + TRANSFORM(m.loException.Details) + "] " + ;
            " (" + TRANSFORM(m.loException.ErrorNo) + ")" + ;
            " in " + m.loException.Procedure + ;
            " on " + TRANSFORM(m.loException.LineNo)

   MESSAGEBOX(m.lcCode, ;
              0+48, ;
              _screen.Caption)
 
ENDTRY

RETURN 

ENDPROC

ENDDEFINE


*: EOF :*  