using System;
using System.Collections.Generic;
using Bytes = ExcelComponent.Excel.ByteUtil.Bytes;

namespace ExcelComponent.Excel
{
    internal class Record
    {
        protected byte[] _rid;
        protected Bytes _data = new Bytes();
        protected List<Record> _continues = new List<Record>();

        public static Record Empty = new Record(Excel.RID.Empty, new byte[0]);

        protected Record()
        {
            //would rather this be a struct, but it's more convenient to be able
            //to use reference equality testing below
        }

        internal Record(byte[] rid, byte[] data) : this(rid, new Bytes(data))
        { }

        internal Record(byte[] rid, Bytes data)
        {
            _rid = Excel.RID.ByteArray(rid);
            int offset = 0;
            int bytesRemaining = data.Length;
            int continueIndex = -1;
            while (bytesRemaining > 0)
            {
                int bytesToAppend = Math.Min(bytesRemaining, BIFF8.MaxDataBytesPerRecord);
                Bytes target;
                if (continueIndex == -1)
                    _data = data.Get(offset, bytesToAppend);
                else
                    _continues.Add(new Record(Excel.RID.CONTINUE, data.Get(offset, bytesToAppend)));
                offset += bytesToAppend;
                bytesRemaining -= bytesToAppend;
                continueIndex++;
            }
        }

        internal byte[] RID
        {
            get { return _rid; }
        }

        internal Bytes Data
        {
            get { return _data; }
        }

        internal List<Record> Continues
        {
            get { return _continues; }
        }

        internal static Bytes GetBytes(byte[] rid, byte[] data)
        {
            return GetBytes(rid, new Bytes(data));
        }

        internal static Bytes GetBytes(byte[] rid, Bytes data)
        {
            if (rid.Length != 2)
                throw new Exception("rid must be 2 bytes ");

            Bytes record = new Bytes();
            
            ushort offset = 0;
            ushort totalLength = (ushort)data.Length;
            do
            {
                ushort length = Math.Min((ushort) (totalLength - offset), BIFF8.MaxDataBytesPerRecord);

                if (offset == 0)
                {
                    record.Append(rid);
                    record.Append(BitConverter.GetBytes(length));
                    record.Append(data.Get(offset, length));
                }
                else
                {
                    record.Append(Excel.RID.CONTINUE);
                    record.Append(BitConverter.GetBytes(length));
                    record.Append(data.Get(offset, length));
                }

                offset += length;
            } while (offset < totalLength);

            return record;
        }

        internal bool IsCellRecord()
        {
            return (_rid == Excel.RID.RK ||
                    _rid == Excel.RID.NUMBER ||
                    _rid == Excel.RID.LABEL ||
                    _rid == Excel.RID.LABELSST ||
                    _rid == Excel.RID.MULBLANK ||
                    _rid == Excel.RID.MULRK ||
                    _rid == Excel.RID.FORMULA);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static List<Record> GetAll(Bytes stream)
        {
            int i = 0;
            List<Record> records = new List<Record>();
            Record lastNonContinue = Record.Empty;
            while (i < (stream.Length - 4))
            {
                byte[] rid = Excel.RID.ByteArray(stream.Get(i, 2).ByteArray);
                Bytes data = new Bytes();
                if (rid == Excel.RID.Empty)
                    break;
                int length = BitConverter.ToUInt16(stream.Get(i + 2, 2).ByteArray, 0);
                data = stream.Get(i + 4, length);
                Record record = new Record(rid, data);
                i += (4 + length);
                if (rid == Excel.RID.CONTINUE)
                {
                    if (lastNonContinue == Record.Empty)
                        throw new Exception("Found CONTINUE record without previous/parent record.");

                    lastNonContinue.Continues.Add(record);
                }
                else
                {
                    lastNonContinue = record;
                    records.Add(record);
                }
            }

            return records;
        }
    }
}
