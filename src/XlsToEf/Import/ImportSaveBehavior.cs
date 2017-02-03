namespace XlsToEf.Import
{
    public class ImportSaveBehavior
    {
        public bool checkForEmptyRows;

        public ImportSaveBehavior()
        {
            RecordMode = RecordMode.Upsert;
            CommitMode = CommitMode.AnySuccessfulAtEndAsBulk;
        }

        public RecordMode RecordMode { get; set; }
        public CommitMode CommitMode { get; set; }
    }
}