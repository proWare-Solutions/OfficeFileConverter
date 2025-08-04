
using OfficeFileConverter.Utils;

namespace OfficeFileConverter.Ppt
{
  internal partial class Binary
  {
    /// <summary>
		/// Record-Header class
		/// 		''' </summary>
		/// 		''' <remarks></remarks>
    internal class RecordHeader
    {
      #region Lokale Eigenschaften
      private ushort _version;
      private ushort _instance;
      private ushort _type;
      private uint _length;
      #endregion
      #region Öffentliche Eigenschaften
      public ushort Version
      {
        get
        {
          return _version;
        }
      }
      #endregion
      public ushort Instance
      {
        get
        {
          return _instance;
        }
      }

      public ushort Type
      {
        get
        {
          return _type;
        }
      }

      public uint Length
      {
        get
        {
          return _length;
        }
      }
      #region Private Methoden

      #endregion
      #region Öffentliche Methoden

      #endregion
      #region Konstruktoren
      internal RecordHeader(ExtMemoryStream mem)
      {
        mem.Read(ref _instance);
        mem.Read(ref _type);
        mem.Read(ref _length);
        // In Instance sind nun die ersten 4 Bits die Version
        _version = (ushort)(_instance & 15);
        _instance = (ushort)(_instance >> 4);
      }
      #endregion

    }
  }
}