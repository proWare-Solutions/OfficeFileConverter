using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using CommandLine;

using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.Word;
using Serilog;
using System.IO;
using Path = System.IO.Path;
using Access = Microsoft.Office.Interop.Access.Application;
using Excel= Microsoft.Office.Interop.Excel.Application;
using PowerPoint = Microsoft.Office.Interop.PowerPoint.Application;
using Visio = Microsoft.Office.Interop.Visio.Application;
using Word = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;
using VisioDoc = Microsoft.Office.Interop.Visio.Document;
using Microsoft.Office.Interop.Access.Dao;
using System.Net.NetworkInformation;

namespace OfficeFileConverter
{
  internal class Program
  {
    private static Options _options;
    private static Access _access;
    private static Excel _excel;
    private static PowerPoint _powerPoint;
    private static Word _word;
    private static Visio _visio;

    static void Main(string[] args)
    {
      DateTime start = DateTime.Now;
      Parser.Default.ParseArguments<Options>(args)
      .WithParsed<Options>(o =>
      {
        Run(o);
      });
      if (_options != null)
      {
        Log.Information($"Total directories accesed: {_options.DirectoriesTotal} files");
        Log.Information($"Total files accesed: {_options.FilesTotal} files");
        Log.Information($"Files processed: {_options.FilesProcessed} files");
        Log.Information($"Duration: {(DateTime.Now-start).ToString(@"d\ hh\:mm\:ss")}");
      }
      Console.WriteLine("Press any key");
      Console.ReadKey();
    }


    static void Run(Options options)
    {
      _options = options;
      LoggerConfiguration conf = new LoggerConfiguration().WriteTo.Console();
      if (options.Debug)
        conf.MinimumLevel.Debug();
      else
        conf.MinimumLevel.Information();
      if (!string.IsNullOrWhiteSpace(options.Logpath))
      {
        options.Logpath = Path.Combine(options.Logpath, options.CurrentRunTime);
        if (!Directory.Exists(options.Logpath)) Directory.CreateDirectory(options.Logpath);
        options.ActionStatusFile = Path.Combine(options.Logpath, "status.csv");
        options.Logpath= System.IO.Path.Combine(options.Logpath, "Scan.log");
        conf.WriteTo.File(options.Logpath + ".log", shared: true);
      }
      Log.Logger = conf.CreateLogger();
      _options.ConvertSettings = new ConvertSettings();
      _options.ConvertSettings.Load(_options.Settings);

      if (!string.IsNullOrWhiteSpace(options.Directory))
      {
        if (!System.IO.Directory.Exists(options.Directory))
        {
          Environment.ExitCode = 1;
          Log.Error($"Directory '{options.Directory}' does not exist");
          return;
        }
        ProcessDirectory(options.Directory);
      }
      else
        if (!string.IsNullOrWhiteSpace(options.File))
        {
          if (!Directory.Exists(options.File))
          {
            Environment.ExitCode = 1;
            Log.Logger.Error($"File '{options.File}' does not exist");
          return;
        }
          ProcessFileList();
        }
      if (_access!=null) _access.Quit();
      if (_excel!=null) _excel.Quit();
      if (_powerPoint!=null) _powerPoint.Quit();
      if (_visio!=null) _visio.Quit();
      if (_word!=null) _word.Quit();
    }

    static void CreateWordApp()
    {
      if (_word == null) 
      { 
        _word = new Word();
        _word.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        _word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
      }
    }

    static void CreateExcelApp()
    {
      if (_excel == null)
      {
        _excel = new Excel();
        _excel.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        _excel.Visible = false;
        _excel.DisplayAlerts = false;
      }
    }

    static void CreatePptApp()
    {
      if (_powerPoint == null)
      {
        _powerPoint = new PowerPoint();
        if (_powerPoint.Visible== Microsoft.Office.Core.MsoTriState.msoTrue) _powerPoint.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        _powerPoint.DisplayAlerts = PpAlertLevel.ppAlertsNone;
        _powerPoint.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
      }
    }

    static void CreateVisioApp()
    {
      if (_visio==null) {
        _visio = new Visio();
        _visio.Visible = false;
      }
    }

    static void CreateAccessApp()
    {
      if (_access == null)
      {
        _access = new Access();
        _access.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        _access.Visible = false;
      }
    }
    static void ProcessFile(string filePath)
    {
      try
      {
        if (!filePath.StartsWith(@"\\?\") && filePath.Length > 250)
        {
          if (filePath.StartsWith(@"\\"))
          {
            // UNC-Paths need UNC-Prefix
            filePath = "UNC" + filePath.Substring(1);
          }
          filePath = @"\\?\" + filePath;
        }
        _options.FilesTotal++;
        ConvertExtension ext = _options.ConvertSettings.GetExtension(filePath);
        if (ext != null && ext.Action == ConvertAction.Convert)
        {
          Log.Information($"Processing file '{filePath}'");
          
          _options.FilesProcessed++;
          string fileExt = ext.Extension.ToLower(); // Endung
          string fileName = Path.GetFileName(filePath); //Dateiname
          string tmpPath = Path.Combine(_options.TempDir, fileName); //Temporärer Pfad
          string fileWihoutExt = Path.GetFileNameWithoutExtension(fileName); //Dateiname ohne Endung
          string tmpTargetPath = tmpTargetPath = Path.Combine(_options.TempDir, fileWihoutExt); // Temporärer Zielpfad ohen Endung
          string targetPath = Path.Combine(Path.GetDirectoryName(filePath), fileWihoutExt);
          Log.Debug("Copying file to temporary directory");
          File.Copy(filePath, tmpPath, true);
          try
          {
            switch (fileExt)
            {
              case "doc":
                {
                  Log.Debug("Processing Word document");
                  CreateWordApp();
                  WordDoc doc = _word.Documents.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      tmpTargetPath += ".pdf";
                      targetPath += ".pdf";
                      doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatPDF);
                     
                    }
                    else
                    {
                      if (doc.HasVBProject)
                      {
                        tmpTargetPath += ".docm";
                        targetPath += ".docm";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatXMLDocumentMacroEnabled);
                      }
                      else
                      {
                        tmpTargetPath += ".docx";
                        targetPath += ".docx";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatXMLDocument);
                      }
                    }
                    try
                    {
                      doc.Close(false);
                      doc = null;
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                      }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally { 
                    if (doc!= null) doc.Close(false); 
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex) 
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}",ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "dot":
                {
                  Log.Debug("Processing Word template");
                  CreateWordApp();
                  WordDoc doc = _word.Documents.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      tmpTargetPath += ".pdf";
                      targetPath += ".pdf";
                      doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatPDF);
                    }
                    else
                    {
                      if (doc.HasVBProject)
                      {
                        tmpTargetPath += ".dotm";
                        targetPath += ".dotm";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatFlatXMLTemplateMacroEnabled);
                      }
                      else
                      {
                        tmpTargetPath += ".dotx";
                        targetPath += ".dotx";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatFlatXMLTemplate);
                      }
                    }
                    try
                    {
                      doc.Close(false);
                      doc = null;
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc!=null) doc.Close(false);
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "xls":
                {
                  Log.Debug("Processing Excel workbook");
                  CreateExcelApp();
                  Workbook wb = _excel.Workbooks.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      List<string> targets = new List<string>();
                      foreach (Worksheet sheet in wb.Worksheets)
                      {
                        if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                        {
                          sheet.Activate();
                          string name = "-" + sheet.Name.Replace('?', '_').Replace('*', '_').Replace('*', '_').Replace('\\', '_').Replace('/', '_') + ".pdf";
                          targets.Add(name);
                          wb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, tmpTargetPath + name);
                        }
                      }
                      foreach (string target in targets)
                      {
                        File.Copy(tmpTargetPath + target, targetPath + target);
                        File.Delete(tmpTargetPath + target);
                      }
                    }
                    else
                    {
                      if (wb.HasVBProject)
                      {
                        tmpTargetPath += ".xlsm";
                        targetPath += ".xlsm";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                      }
                      else
                      {
                        tmpTargetPath += ".xlsx";
                        targetPath += ".xlsx";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLWorkbook);
                      }
                      try
                      {
                        wb.Close(false);
                        wb = null;
                        File.Copy(tmpTargetPath, targetPath, true);
                      }
                      catch (Exception ex)
                      {
                        throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                      }
                      finally
                      {
                        File.Delete(tmpTargetPath);
                      }
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (wb != null) wb.Close(false);
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "xlt":
                {
                  Log.Debug("Processing Excel template");
                  CreateExcelApp();
                  Workbook wb = _excel.Workbooks.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      List<string> targets = new List<string>();
                      foreach (Worksheet sheet in wb.Worksheets)
                      {
                        if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                        {
                          sheet.Activate();
                          string name = "-" + sheet.Name.Replace('?', '_').Replace('*', '_').Replace('*', '_').Replace('\\', '_').Replace('/', '_') + ".pdf";
                          targets.Add(name);
                          wb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, tmpTargetPath + name);
                        }
                      }
                      foreach (string target in targets)
                      {
                        File.Copy(tmpTargetPath + target, targetPath + target);
                        File.Delete(tmpTargetPath + target);
                      }
                    }
                    else
                    {
                      if (wb.HasVBProject)
                      {
                        tmpTargetPath += ".xltm";
                        targetPath += ".xltm";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLTemplateMacroEnabled);
                      }
                      else
                      {
                        tmpTargetPath += ".xltx";
                        targetPath += ".xltx";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLTemplate);
                      }
                      try
                      {
                        wb.Close(false);
                        wb = null;
                        File.Copy(tmpTargetPath, targetPath, true);
                      }
                      catch (Exception ex)
                      {
                        throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                      }
                      finally
                      {
                        File.Delete(tmpTargetPath);
                      }
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (wb != null) wb.Close(false);
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "xla":
                {
                  Log.Debug("Processing Excel addin");
                  CreateExcelApp();
                  Workbook wb = _excel.Workbooks.Open(tmpPath);
                  try
                  {
                    tmpTargetPath += ".xlam";
                    targetPath += ".xlam";
                    wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLAddIn);
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    wb.Close(false);
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "mdb":
                {
                  Log.Debug("Processing Access DB");
                  CreateAccessApp();
                  try
                  {
                    tmpTargetPath += ".accdb";
                    targetPath += ".accdb";
                    _access.ConvertAccessProject(tmpPath, tmpTargetPath, AcFileFormat.acFileFormatAccess2007);
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "vsd":
                {
                  Log.Debug("Processing Visio drawing");
                  CreateVisioApp();
                  VisioDoc doc= _visio.Documents.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      tmpTargetPath += ".pdf";
                      targetPath += ".pdf";
                      doc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, tmpTargetPath, VisDocExIntent.visDocExIntentPrint, VisPrintOutRange.visPrintAll);
                    }
                    else
                    {
                      if (doc.VBProjectData.Length > 0)
                      {
                        tmpTargetPath += ".vsdm";
                        doc.Version = VisDocVersions.visVersion140;
                        targetPath += ".vsdm";
                      }
                      else
                      {
                        tmpTargetPath += ".vsdx";
                        targetPath += ".vsdx";
                        doc.Version = VisDocVersions.visVersion140;
                      }
                      doc.SaveAs(tmpTargetPath);
                    }
                    doc.Close();
                    doc = null;
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc != null) doc.Close();
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "vst":
                {
                  Log.Debug("Processing Visio template");
                  CreateVisioApp();
                  VisioDoc doc = _visio.Documents.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      tmpTargetPath += ".pdf";
                      targetPath += ".pdf";
                      doc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, tmpTargetPath, VisDocExIntent.visDocExIntentPrint, VisPrintOutRange.visPrintAll);
                    }
                    else
                    {
                      if (doc.VBProjectData.Length > 0)
                      {
                        tmpTargetPath += ".vstm";
                        targetPath += ".vstm";
                        doc.Version = VisDocVersions.visVersion140;
                      }
                      else
                      {
                        tmpTargetPath += ".vstx";
                        targetPath += ".vstx";
                        doc.Version = VisDocVersions.visVersion140;
                      }
                      doc.SaveAs(tmpTargetPath);
                    }
                    doc.Close();
                    doc = null;
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc != null) doc.Close();
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "vss":
                {
                  Log.Debug("Processing Visio drawing");
                  CreateVisioApp();
                  VisioDoc doc = _visio.Documents.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      tmpTargetPath += ".pdf";
                      targetPath += ".pdf";
                      doc.ExportAsFixedFormat(VisFixedFormatTypes.visFixedFormatPDF, tmpTargetPath, VisDocExIntent.visDocExIntentPrint, VisPrintOutRange.visPrintAll);
                    }
                    else
                    {
                      if (doc.VBProjectData.Length > 0)
                      {
                        tmpTargetPath += ".vssm";
                        targetPath += ".vssm";
                        doc.Version = VisDocVersions.visVersion140;
                      }
                      else
                      {
                        tmpTargetPath += ".vssx";
                        targetPath += ".vssx";
                        doc.Version = VisDocVersions.visVersion140;
                      }
                      doc.SaveAs(tmpTargetPath);
                    }
                    doc.Close();
                    doc = null;
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc != null) doc.Close();
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "ppt":
                {
                  Log.Debug("Processing PowerPoint presentation");
                  CreatePptApp();
                  Presentation presentation = _powerPoint.Presentations.Open(tmpPath,Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      tmpTargetPath += ".pdf";
                      targetPath += ".pdf";
                      presentation.SaveCopyAs(tmpTargetPath, PpSaveAsFileType.ppSaveAsPDF);
                    }
                    else
                    {
                      if (presentation.HasVBProject)
                      {
                        targetPath += ".pptm";
                        tmpTargetPath += ".pptm";
                      }
                      else
                      {
                        targetPath += ".pptx";
                        tmpTargetPath += ".pptx";
                      }
                      presentation.SaveAs(tmpTargetPath);
                    }
                    presentation.Close();
                    presentation = null;
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (presentation != null) presentation.Close();
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "pot":
                {
                  Log.Debug("Processing PowerPoint template");
                  CreatePptApp();
                  Presentation presentation = _powerPoint.Presentations.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      tmpTargetPath += ".pdf";
                      targetPath += ".pdf";
                      presentation.SaveCopyAs(tmpTargetPath, PpSaveAsFileType.ppSaveAsPDF);
                    }
                    else
                    {
                      if (presentation.HasVBProject)
                      {
                        targetPath += ".potm";
                        tmpTargetPath += ".potm";
                      }
                      else
                      {
                        targetPath += ".potx";
                        tmpTargetPath += ".potx";
                      }
                      presentation.SaveAs(tmpTargetPath);
                    }
                    presentation.Close();
                    presentation = null;
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (presentation != null) presentation.Close();
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
              case "pps":
                {
                  Log.Debug("Processing PowerPoint show");
                  CreatePptApp();
                  Presentation presentation = _powerPoint.Presentations.Open(tmpPath);
                  try
                  {
                    if (ext.TargetIsPDF)
                    {
                      targetPath += ".pdf";
                      tmpTargetPath += ".pdf";
                      presentation.SaveCopyAs(tmpTargetPath, PpSaveAsFileType.ppSaveAsPDF);
                    }
                    else
                    {
                      if (presentation.HasVBProject)
                      {
                        targetPath += ".ppsm";
                        tmpTargetPath += ".ppsm";
                      }
                      else
                      {
                        targetPath += ".ppsx";
                        tmpTargetPath += ".ppsx";
                      }
                      presentation.SaveAs(tmpTargetPath);
                    }
                    presentation.Close();
                    presentation = null;
                    try
                    {
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (presentation != null) presentation.Close();
                  }
                  if (_options.RemoveOriginal)
                  {
                    try
                    {
                      File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                      throw new ConvertException($"Could not remove original file: {ex.Message}", ex, ActionStatus.OriginalRemoveFailed);
                    }
                  }
                  break;
                }
            }
          }
          catch (ConvertException ex)
          {
            if (!string.IsNullOrEmpty(_options.ActionStatusFile)) File.AppendAllText(_options.ActionStatusFile, $"File {fileName} process with result {ex.Status}" + Environment.NewLine);
          }
          catch
          {
            if (!string.IsNullOrEmpty(_options.ActionStatusFile)) File.AppendAllText(_options.ActionStatusFile, $"File {fileName} process with result {ActionStatus.Unknown}" + Environment.NewLine);
          }
          finally 
          {
            if (File.Exists(tmpPath))
            {
              Log.Debug("Removing temporary file");
              File.Delete(tmpPath);
            }
          }
        }
      }
      catch (Exception ex)
      {
        Log.Error($"Error processing file '{filePath}': {ex.Message}");
      }

    }

    static void ProcessDirectory(string directory)
    {
      try
      {
        Log.Information($"Processing directory '{directory}'");
        _options.DirectoriesTotal++;
        foreach (string file in Directory.GetFiles(directory))
        {
          ProcessFile(file);
        }
        foreach (string subDir in Directory.GetDirectories(directory))
        {
          try
          {
            ProcessDirectory(subDir);
          }
          catch (PathTooLongException ex)
          {
            if (!directory.StartsWith(@"\\?\") && directory.Length > 250)
            {
              if (directory.StartsWith(@"\\"))
              {
                // UNC-Paths need UNC-Prefix
                directory = "UNC" + directory.Substring(1);
              }
              ProcessDirectory(@"\\?\" + directory);
            }
            else
            {
              Log.Warning(string.Format("Path too long {0}: {1}", directory, ex.Message));
            }
          }
          catch (DirectoryNotFoundException ex)
          {
            try
            {
              if (!directory.StartsWith(@"\\?\") && directory.Length > 250)
              {
                if (directory.StartsWith(@"\\"))
                {
                  // UNC-Paths need UNC-Prefix
                  directory = "UNC" + directory.Substring(1);
                }
                ProcessDirectory(@"\\?\" + directory);
              }
              else
              {
                Log.Warning(string.Format("Directory not found {0}: {1}", directory, ex.Message));
              }
            }
            catch (System.IO.DirectoryNotFoundException innerEx)
            {
              Log.Warning(string.Format("Directory not found {0}: {1}", directory, innerEx.Message));
            }
          }
          catch (Exception)
          {
            throw;
          }
        }
      }
      catch (Exception ex)
      {
        Log.Error($"Error processing directory '{directory}': {ex.Message}");
      }
    }

    static void ProcessFileList()
    {
      foreach (string file in File.ReadLines(_options.File))
      {
        if (!string.IsNullOrWhiteSpace(file)) ProcessFile(file);
      }
    }
  }
}
