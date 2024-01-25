using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeFileConverter
{
  /// <summary>
  /// Klass mit Settings zum Konvertiren von Dateien
  /// </summary>

  public enum ConvertAction
  {
    Convert,
    Skip
  }
  [Serializable]
  public class ConvertExtension
  {
    public string Extension { get; set; }

    public bool TargetIsPDF { get; set; } = false;
    public ConvertAction Action { get; set; } = ConvertAction.Convert;
    public DateTime ValidFrom { get; set; }
  }
  [Serializable]
  public class ConvertSettings
  {
    [NonSerialized]
    private Dictionary<string, SortedList<DateTime, ConvertExtension>> _settingsDic = new Dictionary<string, SortedList<DateTime, ConvertExtension>>();

    public List<ConvertExtension> Settings { get; set; } = new List<ConvertExtension>();

    public void AddSetting(ConvertExtension setting)
    {
      Settings.Add(setting);
      SortedList<DateTime, ConvertExtension> list = new SortedList<DateTime, ConvertExtension>();
      if (_settingsDic.ContainsKey(setting.Extension.ToLower()))
        list = _settingsDic[setting.Extension.ToLower()];
      else
        _settingsDic.Add(setting.Extension.ToLower(), list);
      list.Add(setting.ValidFrom, setting);
    }
    public void Load(string path)
    {
      XmlDocument doc = new XmlDocument();
      try
      {
        doc.Load(path);
        Load(doc);
      }
      catch
      {
        throw new ApplicationException("Settings Datei kann nicht gelesen werden");
      }
    }
    public void Load(System.Xml.XmlDocument xmlDoc)
    {
      try
      {
        System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(ConvertSettings));
        using (System.Xml.XmlNodeReader reader = new System.Xml.XmlNodeReader(xmlDoc.DocumentElement))
        {
          Load((ConvertSettings)serializer.Deserialize(reader));
        }
      }
      catch 
      {
        throw new ApplicationException("Fehlerhafte Settings-Datei");
      }
    }

    internal void Save(string path)
    {
      System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(ConvertSettings));
      using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(path))
      {
        serializer.Serialize(writer, this);
      }
    }

    public void Load(ConvertSettings settings)
    {
      if (settings == null)
        throw new ArgumentException("settings");
      foreach (ConvertExtension ext in settings.Settings)
      {
        SortedList<DateTime, ConvertExtension> extensions = new SortedList<DateTime, ConvertExtension>();
        if (_settingsDic.ContainsKey(ext.Extension.ToLower()))
          extensions = _settingsDic[ext.Extension.ToLower()];
        else
          _settingsDic.Add(ext.Extension.ToLower(), extensions);
        extensions.Add(ext.ValidFrom, ext);
      }
    }
    /// <summary>
    /// Liefert ein Extension-Objekt zu einem Dateinamen und einem Dateidatum
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="validFrom"></param>
    /// <returns></returns>
    public ConvertExtension GetExtension(string filePath)
    {
      string extension = System.IO.Path.GetExtension(filePath);
      if (extension.StartsWith(".")) { extension = extension.Substring(1); }
      extension = extension.ToLower();
      DateTime validFrom = System.IO.File.GetLastWriteTime(filePath);
      ConvertExtension result = null;
      if (_settingsDic.ContainsKey(extension))
      {
        SortedList<DateTime, ConvertExtension> list = _settingsDic[extension];
        DateTime foundKey = DateTime.MinValue;
        foreach (DateTime vf in list.Keys)
        {
          if (validFrom >= vf)
            foundKey = vf;
        }
        if (foundKey != DateTime.MinValue)
          result = list[foundKey];
      }
      return result;
    }
  }
}
