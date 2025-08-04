using OpenMcdf;
using OfficeFileConverter.Utils;
using System.Collections.Generic;
using System;

namespace OfficeFileConverter.Ppt
{
  /// <summary>
	/// Powerpoint binary class
	/// 	''' </summary>
	/// 	''' <remarks></remarks>
  internal partial class Binary 
  {

    #region Lokale Eigenschaften
    private CFStorage _storage = null;
    private byte[] _vbaStorage = null;
    private Dictionary<uint, PersistDirEntry> _persistenObjectsDirectory = new Dictionary<uint, PersistDirEntry>();
    private bool _encrypted = false;
    #endregion
    #region Öffentliche Eigenschaften
    public byte[] VBAStorage
    {
      get
      {
        return _vbaStorage;
      }
    }

    /// <summary>
    /// True, wenn das Dokument verschlüsselt ist
    /// </summary>
    /// <value></value>
    /// <returns></returns>
    /// <remarks></remarks>
    public bool Encrypted
    {
      get
      {
        return _encrypted;
      }
    }
    #endregion
    #region Private Methoden
   

    /// <summary>
		/// 		''' Liest den Versatz zum UserEditAtom
		/// 		''' </summary>
		/// 		''' <returns></returns>
		/// 		''' <remarks></remarks>
    private uint GetOffsetToUserEdit()
    {
      uint offsetToCurrentEdit = 0U;
      var userStream = _storage.GetStream("Current User");
      using (var mem = new ExtMemoryStream(userStream.GetData()))
      {
        var header = new RecordHeader(mem);
        if (header.Type != 0xFF6)
          throw new ApplicationException("Error reading offset to UserEditAtom");
        uint ui = 0U;
        mem.Read(ref ui); // Size
        if (ui != 0x14L)
          throw new ApplicationException("Error reading offset to UserEditAtom");
        mem.Read(ref ui); // HeaderToken
        if (ui == 4090610911L)
          _encrypted = true; // Hex-Wert F3D1C4DF
        mem.Read(ref offsetToCurrentEdit);
        mem.Close();
      }



      return offsetToCurrentEdit;
    }

    #endregion

    #region Konstruktoren
    internal Binary(CompoundFile cf)
    {
      _storage = cf.RootStorage;
      uint offsetToUserEdit = GetOffsetToUserEdit();
      if (offsetToUserEdit == 0L)
        throw new ApplicationException("Invalid offset to UserEditAtom");
      
    }
    #endregion

  }
}