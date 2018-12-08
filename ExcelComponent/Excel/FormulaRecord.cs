namespace ExcelComponent.Excel
{
    internal class FormulaRecord : Record
    {
        internal Record StringRecord = null;

        internal FormulaRecord(Record formulaRecord, Record stringRecord)
            : base()
        {
            _rid = formulaRecord.RID;
            _data = formulaRecord.Data;
            _continues = formulaRecord.Continues;

            StringRecord = stringRecord;
        }
    }
}
