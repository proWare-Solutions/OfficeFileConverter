using System;
using System.IO;

namespace OfficeFileConverter.Utils
{
  internal class ExtMemoryStream : MemoryStream
  {

    #region Emums
    public enum Endian : byte
    {
      LittleEndian = 0,
      BigEndian = 1
    }
    #endregion

    #region Lokale Eigenschaften
    private Endian _numberFormat = Endian.LittleEndian;
    #endregion

    #region Öffentliche Eigenschaften
    public Endian NumberFormat
    {
      set
      {
        _numberFormat = value;
      }
    }
    #endregion

    #region Öffentliche Methoden
    public void Read(ref ushort intRef, Endian endian)
    {
      var buf = new byte[2];
      Read(buf, 0, 2);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      intRef = BitConverter.ToUInt16(buf, 0);
    }

    public void Read(ref ushort intRef)
    {
      Read(ref intRef, _numberFormat);
    }

    public void Read(ref short intRef)
    {
      var buf = new byte[2];
      Read(buf, 0, 2);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      intRef = BitConverter.ToInt16(buf, 0);
    }

    public void Read(ref uint intRef)
    {
      var buf = new byte[4];
      Read(buf, 0, 4);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      intRef = BitConverter.ToUInt32(buf, 0);
    }

    public void Read(ref ulong intRef)
    {
      var buf = new byte[8];
      Read(buf, 0, 8);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      intRef = BitConverter.ToUInt32(buf, 0);
    }

    public void Read(ref int intRef)
    {
      var buf = new byte[4];
      Read(buf, 0, 4);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      intRef = BitConverter.ToInt32(buf, 0);
    }

    public void Write(ushort intRef, Endian endian)
    {
      var buf = new byte[2];
      buf = BitConverter.GetBytes(intRef);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      Write(buf, 0, 2);
    }

    public void Write(ushort intRef)
    {
      Write(intRef, _numberFormat);
    }

    public void Write(uint intRef)
    {
      var buf = BitConverter.GetBytes(intRef);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      Write(buf, 0, 4);
    }

    public void Write(ulong intRef)
    {
      var buf = BitConverter.GetBytes(intRef);
      if (BitConverter.IsLittleEndian && _numberFormat == Endian.BigEndian)
        Array.Reverse(buf);
      Write(buf, 0, 8);
    }
    #endregion
    #region Konstruktoren
    public ExtMemoryStream() : base()
    {
    }

    public ExtMemoryStream(byte[] buf) : base(buf)
    {
    }

    public ExtMemoryStream(int capacity) : base(capacity)
    {
    }
    #endregion
  }
}