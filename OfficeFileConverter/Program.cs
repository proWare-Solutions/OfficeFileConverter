using CommandLine;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.Word;
using OfficeFileConverter.Ppt;
using OpenMcdf;
using OpenMcdf.Extensions;
using OpenMcdf.Extensions.OLEProperties;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using Access = Microsoft.Office.Interop.Access.Application;
using Excel= Microsoft.Office.Interop.Excel.Application;
using Path = System.IO.Path;
using PowerPoint = Microsoft.Office.Interop.PowerPoint.Application;
using Visio = Microsoft.Office.Interop.Visio.Application;
using VisioDoc = Microsoft.Office.Interop.Visio.Document;
using Word = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;


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

    private static StringBuilder _filesInUse = new StringBuilder();
    private static StringBuilder _filesEncrypted = new StringBuilder();
    private static StringBuilder _filesFailed = new StringBuilder();
    private static StringBuilder _filesNotConverted = new StringBuilder();

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
        string log = System.IO.Path.Combine(options.Logpath, "Scan.log");
        conf.WriteTo.File(log, shared: true);
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
      if (!string.IsNullOrEmpty(_options.ActionStatusFile)) File.WriteAllText(_options.ActionStatusFile, _filesNotConverted.ToString());
      if (!string.IsNullOrWhiteSpace(options.Logpath))
      {
        if (_filesInUse.Length > 0) File.WriteAllText(Path.Combine(options.Logpath, "FilesInUse.txt"), _filesInUse.ToString());
        if (_filesEncrypted.Length > 0) File.WriteAllText(Path.Combine(options.Logpath, "FilesEncrypted.txt"), _filesEncrypted.ToString());
        if (_filesFailed.Length > 0) File.WriteAllText(Path.Combine(options.Logpath, "FilesFailed.txt"), _filesFailed.ToString());
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

    static bool IsFileLocked(string path)
    {
      try
      {
        using (FileStream stream = new FileStream(path,FileMode.Open, FileAccess.Read, FileShare.None))
        {
          stream.Close();
        }
      }
      catch (IOException)
      {
        //the file is unavailable because it is:
        //still being written to
        //or being processed by another thread
        //or does not exist (has already been processed)
        return true;
      }

      //file is not locked
      return false;
    }

    static bool IsWordOrExcelEncrypted(string path)
    {
      // Open structured storage files
      using (CompoundFile cf = new CompoundFile(path, CFSUpdateMode.ReadOnly, CFSConfiguration.Default))
      {
        CFStream ds;
        if (cf.RootStorage.TryGetStream("\u0005SummaryInformation", out ds))
        {
          OLEPropertiesContainer props = ds.AsOLEPropertiesContainer();
          foreach (var prop in props.Properties)
          {
            if (prop.PropertyIdentifier == 19)
            {
              return (int)prop.Value == 1;
            }
          }
        }
        
      }
      return false;
    }

    static bool IsPowerPointEncrypted(string path)
    {
      using (CompoundFile cf = new CompoundFile(path, CFSUpdateMode.ReadOnly, CFSConfiguration.Default))
      {
        Binary binary = new Binary(cf);
        return binary.Encrypted;
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
          if (IsFileLocked(filePath))
          {
            Log.Information($"File {filePath} is in use and cannot be processed");
            _filesInUse.AppendLine(filePath);
            return;
          }
          _options.FilesProcessed++;
          string fileExt = ext.Extension.ToLower(); // Endung
          string fileName = Path.GetFileName(filePath); //Dateiname
          if (fileName.StartsWith("~$")) return; // Lock-File for a file in use
          string tmpPath = Path.Combine(_options.TempDir, fileName); //Temporärer Pfad
          string fileWihoutExt = Path.GetFileNameWithoutExtension(fileName); //Dateiname ohne Endung
          string tmpTargetPath = tmpTargetPath = Path.Combine(_options.TempDir, fileWihoutExt); // Temporärer Zielpfad ohen Endung
          string targetPath = Path.Combine(Path.GetDirectoryName(filePath), fileWihoutExt);
          Log.Debug("Copying file to temporary directory");
          File.Copy(filePath, tmpPath, true);
          bool fileChanged = false;
          try
          {
            switch (fileExt)
            {
              case "doc":
                {
                  if (IsWordOrExcelEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                      fileChanged = true;
                    }
                    else
                    {
                      if (doc.HasVBProject)
                      {
                        tmpTargetPath += ".docm";
                        targetPath += ".docm";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatXMLDocumentMacroEnabled);
                        fileChanged = true;
                      }
                      else if (_options.AllFiles)
                      {
                        tmpTargetPath += ".docx";
                        targetPath += ".docx";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatXMLDocument);
                        fileChanged = true;
                      }
                    }
                    try
                    {
                      doc.Close(false);
                      doc = null;
                      if (fileChanged && !ext.TargetIsPDF) File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      fileChanged = false;
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {

                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally 
                  { 
                    if (doc!= null) doc.Close(false); 
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
                  if (IsWordOrExcelEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                      fileChanged = true;
                    }
                    else
                    {
                      if (doc.HasVBProject)
                      {
                        tmpTargetPath += ".dotm";
                        targetPath += ".dotm";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatFlatXMLTemplateMacroEnabled);
                        fileChanged = true;
                      }
                      else if (_options.AllFiles)
                      {
                        tmpTargetPath += ".dotx";
                        targetPath += ".dotx";
                        doc.SaveAs2(tmpTargetPath, WdSaveFormat.wdFormatFlatXMLTemplate);
                        fileChanged = true;
                      }
                    }
                    try
                    {
                      doc.Close(false);
                      doc = null;
                      if (fileChanged && !ext.TargetIsPDF) File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      fileChanged = false;
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc!=null) doc.Close(false);
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
                  if (IsWordOrExcelEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                      fileChanged = true;
                    }
                    else
                    {
                      if (wb.HasVBProject)
                      {
                        tmpTargetPath += ".xlsm";
                        targetPath += ".xlsm";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                        fileChanged = true;
                      }
                      else if (_options.AllFiles)
                      {
                        tmpTargetPath += ".xlsx";
                        targetPath += ".xlsx";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLWorkbook);
                        fileChanged = true;
                      }
                      try
                      {
                        wb.Close(false);
                        wb = null;
                        if (fileChanged && !ext.TargetIsPDF) File.Copy(tmpTargetPath, targetPath, true);
                      }
                      catch (Exception ex)
                      {
                        fileChanged = false;
                        throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                      }
                      finally
                      {
                        File.Delete(tmpTargetPath);
                      }
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (wb != null) wb.Close(false);
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
                  if (IsWordOrExcelEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                      fileChanged = true;
                    }
                    else
                    {
                      if (wb.HasVBProject)
                      {
                        tmpTargetPath += ".xltm";
                        targetPath += ".xltm";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLTemplateMacroEnabled);
                        fileChanged = true;
                      }
                      else
                      {
                        tmpTargetPath += ".xltx";
                        targetPath += ".xltx";
                        wb.SaveAs(tmpTargetPath, XlFileFormat.xlOpenXMLTemplate);
                        fileChanged = true;
                      }
                      try
                      {
                        wb.Close(false);
                        wb = null;
                        if (fileChanged && !ext.TargetIsPDF) File.Copy(tmpTargetPath, targetPath, true);
                      }
                      catch (Exception ex)
                      {
                        fileChanged = false;
                        throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                      }
                      finally
                      {
                        File.Delete(tmpTargetPath);
                      }
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (wb != null) wb.Close(false);
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
                  if (IsWordOrExcelEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                    fileChanged = false;
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    wb.Close(false);
                    try
                    {
                      if (fileChanged)
                        File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      fileChanged = false;
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
                  if (!_options.ConvertAccess) return;
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
                      fileChanged = true;
                    }
                    catch (Exception ex)
                    {
                      fileChanged = false;
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
              case "vdx":
              case "vsd":
              case "vdw":
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
                      fileChanged = true;
                    }
                    catch (Exception ex)
                    {
                      fileChanged = false;
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc != null) doc.Close();
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
              case "vtx":
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
                      fileChanged = true;
                      File.Copy(tmpTargetPath, targetPath, true);
                    }
                    catch (Exception ex)
                    {
                      fileChanged = false;
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc != null) doc.Close();
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
              case "vsx":
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
                      fileChanged = true;
                    }
                    catch (Exception ex)
                    {
                      fileChanged = false;
                      throw new ConvertException($"Could not save new file {targetPath}: {ex.Message}", ex, ActionStatus.SaveFailed);
                    }
                    finally
                    {
                      File.Delete(tmpTargetPath);
                    }
                  }
                  catch (ConvertException)
                  {
                    throw;
                  }
                  catch (Exception ex)
                  {
                    throw new ConvertException($"Could not save new file '{tmpTargetPath}': {ex.Message}", ex, ActionStatus.SaveFailed);
                  }
                  finally
                  {
                    if (doc != null) doc.Close();
                  }
                  if (_options.RemoveOriginal && fileChanged)
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
              case "ppa":
                {
                  if (IsPowerPointEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
                  Log.Debug("Processing PowerPoint Add-In");
                  CreatePptApp();
                  Presentation presentation = _powerPoint.Presentations.Open(tmpPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
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
                        targetPath += ".ppam";
                        tmpTargetPath += ".ppam";
                      }
                      else
                      {
                        targetPath += ".ppax";
                        tmpTargetPath += ".ppax";
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
                  catch (ConvertException)
                  {
                    throw;
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
              case "ppt":
                {
                  if (IsPowerPointEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                  catch (ConvertException)
                  {
                    throw;
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
                  if (IsPowerPointEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                  catch (ConvertException) 
                  { 
                    throw;
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
                  if (IsPowerPointEncrypted(tmpPath))
                  {
                    Log.Information($"File {fileName} is encrypted and cannot be converted");
                    _filesEncrypted.AppendLine(filePath);
                    return;
                  }
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
                  catch (ConvertException)
                  {
                    throw;
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
            _filesNotConverted.AppendLine($"File {filePath} process with result {ex.Status}");
            _filesFailed.AppendLine(filePath);
          }
          catch
          {
            _filesNotConverted.AppendLine($"File {filePath} process with result {ActionStatus.Unknown}");
            _filesFailed.AppendLine(filePath);
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
        _filesFailed.AppendLine(filePath);
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
