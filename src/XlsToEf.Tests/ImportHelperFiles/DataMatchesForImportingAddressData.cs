using MediatR;
using XlsToEf.Import;

namespace XlsToEf.Tests.ImportHelperFiles
{
    public class DataMatchesForImportingAddressData : DataMatchesForImport, IAsyncRequest<ImportResult>
    {
    }
}