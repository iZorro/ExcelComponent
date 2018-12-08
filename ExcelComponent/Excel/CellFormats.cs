using System.Collections.Generic;
using Bytes = ExcelComponent.Excel.ByteUtil.Bytes;

namespace ExcelComponent.Excel
{
    /// <summary>
    /// Manages the collection of XF objects for a Workbook.
    /// </summary>
	public class CellFormats
	{
		private readonly XlsDocument _doc;
		private readonly Workbook _workbook;

        private CellFormat _defaultUserXf;

		private readonly List<CellFormat> _xfs;

        internal SortedList<ushort, ushort> XfIdxLookups = null;

	    internal CellFormats(XlsDocument doc, Workbook workbook)
		{
			_doc = doc;
			_workbook = workbook;

			_xfs = new List<CellFormat>();

			AddDefaultStyleXFs();
			AddDefaultUserXF();
	        //AddDefaultFormattedStyleXFs();  //what was I thinking about here?
		}

        internal CellFormat this[int index]
        {
            get { return (CellFormat)_xfs[index].Clone(); }
        }

        internal CellFormat DefaultUserXF
        {
            get { return (CellFormat)_defaultUserXf.Clone(); }
        }

        internal ushort Add(CellFormat xf)
		{
            _workbook.Fonts.Add(xf.Font);
            _workbook.Formats.Add(xf.Format);
            _workbook.Styles.Add(xf.Style);

            //TODO: What happens if they try to re-add a Default (i.e. non-user) XF?
			short xfId = GetId(xf);
			if (xfId == -1)
			{
				xfId = (short)_xfs.Count;
				_xfs.Add((CellFormat)xf.Clone());
			}

			//NOTE: Not documented, but User-defined XFs must have a minimum
			//index of 16 (0-based).

			return (ushort)xfId;
		}

		private void AddDefaultStyleXFs()
		{
		    CellFormat xf = new CellFormat(_doc);
			xf.IsStyleXF = true;
			xf.CellLocked = true; //TODO: Is this correct?  Default Style XF is CellLocked?  what's the origin of this line?
		    _xfs.Add(xf);

		    xf = (CellFormat) xf.Clone();
            xf.UseBackground = false;
            xf.UseBorder = false;
            xf.UseFont = true;
            xf.UseMisc = false;
            xf.UseNumber = false;
            xf.UseProtection = false;

            //Gotta have a 16th (index 15) XF for the Default Cell Format...
            //See excelfileformat.pdf Sec. 4.6.2: The default cell format is always present
            //in an Excel file, described by the XF record with the fixed index 15 (0-based).
            //By default, it uses the worksheet/workbook default cell style, described by
            //the very first XF record (index 0);
            //Apparently Excel 2003 was okay without it, but 2007 chokes if it's not there.
            for (int i = 0; i < 15; i++)
    		    _xfs.Add((CellFormat) xf.Clone());
		}

		private void AddDefaultUserXF()
		{
			CellFormat xf = new CellFormat(_doc);
		    xf.CellLocked = true;

            Add(xf);

            _defaultUserXf = xf;
        }

//        private void AddDefaultFormattedStyleXFs()
//        {
//            
//        }

		private short GetId(CellFormat xf)
		{
			for (short i = 0; i < _xfs.Count; i++)
				if (_xfs[i].Equals(xf))
					return i;

			return -1;
		}

	    internal Bytes Bytes
		{
			get
			{
				Bytes bytes = new Bytes();

				for (int i = 0; i < _xfs.Count; i++)
				{
					bytes.Append(_xfs[i].Bytes);
				}

				return bytes;
			}
		}

        /// <summary>
        /// Gets the number of XF objects in this XFs collection.
        /// </summary>
        public object Count
        {
            get { return _xfs.Count; }
        }
	}
}
