******************************************************************************************
*  PROGRAM: Thor_Tool_ReportLabelPrinterCleanserSingleFile.prg
*
*  AUTHOR: Richard A. Schummer, March 2015
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
*     This program removes the rarely useful hardcoded developer printer details from 
*     reports and labels in a single file. You are prompted to pick the folder where the reports  
*     reside. Next, you are asked if you want to review the contents of the report columns that 
*     store the printer details in the scrubber log.report scrubber log.
*
*  CALLING SYNTAX:
*     DO Thor_Tool_ReportLabelPrinterCleanser.prg
*     Thor_Tool_ReportLabelPrinterCleanser()
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
* 03/14/2015  Richard A. Schummer     1.0      Created Program
* ----------------------------------------------------------------------------------------
* 09/20/2015  Richard A. Schummer     1.1      Updated Thor registration information and 
*                                              general code cleanup
* ----------------------------------------------------------------------------------------
*
******************************************************************************************
LPARAMETERS lxParam1

#DEFINE ccCRLF            CHR(13)+CHR(10)
#DEFINE ccTOOLNAME        "WLC Report/Label Printer Detail Scrubber Tool"

* Standard prefix for all tools for Thor, allowing this tool to tell Thor about itself.
IF PCOUNT() = 1 AND 'O' = VARTYPE(lxParam1) AND 'thorinfo' == LOWER(lxParam1.Class)
   WITH lxParam1
   
      * Required
      .Prompt          = "Report/Label Printer Detail Scrubber Tool - Single File"           && used in menus
      .StatusBarText   = "Remove the rarely useful hardcoded developer printer details from a report or label."  
      
      * Optional
      TEXT TO .Description NOSHOW PRETEXT 1+2     && a description for the tool
         This program removes the rarely useful hardcoded developer printer details from a report and label, one file at a time.      
      ENDTEXT  
      
      .CanRunAtStartUp = .F.

      * These are used to group and sort tools when they are displayed in menus or the Thor form
      .Source        = "WLC"                      && where did this tool come from?  Your own initials, for instance
      .Category      = "WLC"                      && creates categorization of tools; defaults to .Source if empty
      .Sort          = 0                          && the sort order for all items from the same Category
      
      * For public tools, such as PEM Editor, etc.
      .Version       = "Version 2.2, September 20, 2015"           && e.g., 'Version 7, May 18, 2011'
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
*    Runs the main tool code.
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

#DEFINE ccCLASS     "cushookcleanreports"
#DEFINE ccCLASSLIB  "d:\devvfp8apps\devtools\projecthook\cprojecthook5.vcx"
#DEFINE ccLOGFILE   "WLCScrubReportPrinters.txt"
#DEFINE ccCRLF      CHR(13)+CHR(10)


LOCAL loException as Exception, ;
      lcOldSafety, ;
      lcCode, ;
      lcLogText, ;
      llViewContents, ;
      lnFiles, ;
      lnResponse, ;
      lnSorted
      
lcOldSafety = SET("Safety")
SET SAFETY OFF

* Have developer choose the reports folder
WAIT WINDOW "Select a report or label..." NOWAIT NOCLEAR 
lcFileName = GETFILE("frx;lbx")
WAIT CLEAR 

IF EMPTY(lcFileName)
   * Nothing to do
ELSE 
   * Have developer decide level of logging with content of the Tag, Tag2, and Expr column if data exists
   lnResponse = MESSAGEBOX("Do you want to view the contents of the report printer settings columns in the log?", ;
                           0+4+32+256, ;
                           _screen.Caption)

   llViewContents = lnResponse = 6

   TRY
      lcLogText = SPACE(0)

      lcLogText = m.lcLogText + ;
                  ccTOOLNAME + ccCRLF + ccCRLF + ;
                  TRANSFORM(DATETIME()) + ccCRLF + ccCRLF  

      * Process a single file
      ScrubPrinterDetails(lcFileName, @lcLogText, llViewContents)

      * Output log
      STRTOFILE(m.lcLogText, ccLOGFILE, 0)
      MODIFY FILE (ccLOGFILE) NOEDIT NOWAIT RANGE 1,1
      
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
ENDIF 

SET SAFETY &lcOldSafety

RETURN  


********************************************************************************
*  METHOD NAME: ScrubPrinterDetails
*
*  AUTHOR: Richard A. Schummer, March 2015
*
*  METHOD DESCRIPTION:
*    Process that loops through the files and calls the routine to optionally
*    view and scrub the hardcoded printer details.
*
*  INPUT PARAMETERS:
*    tcFileName     = character, required, the name of the report/label filename
*    tcLogText      = character passed by reference, required, log text passed in and updated.
*    tlViewContents = logical, determines of the contents of the FRX/LBX file are viewed
* 
*  OUTPUT PARAMETERS:
*    None
* 
********************************************************************************
PROCEDURE ScrubPrinterDetails(tcFileName, tcLogText, tlViewContents)

LOCAL lcCode, ;
      lcOneFileLog, ;
      lnFiles, ;
      lnFileNameLength, ;
      lnI, ;
      loException AS Exception, ;
      loReports

TRY 
   loReports = NEWOBJECT(ccCLASS, ccCLASSLIB)
   loReports.lIncludeHeading = .F.
   
   * First record the contents if requested
   IF tlViewContents
      loReports.Clean(tcFileName, "view")

      tcLogText = tcLogText + ;
                  loReports.GetIssueLog() + ;
                  ccCRLF
   ELSE
      tcLogText = tcLogText + ;
                  tcFileName
   ENDIF 
   
   * Then do the actual scrubbing of the printer information from file
   loReports.Clean(tcFileName, "clean")
   lcOneFileLog = loReports.GetIssueLog()

   tcLogText = tcLogText + ;
               IIF(EMPTY(lcOneFileLog), " - Nothing to clean.", ccCRLF + lcOneFileLog + ccCRLF) + ;
               ccCRLF

   loReports.Release()

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


*: EOF :*  