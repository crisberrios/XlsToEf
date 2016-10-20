using MediatR;
using XlsToEf.Import;

namespace XlsToEf.Example.ExampleCustomMapperField
{
    public class XlsProductColumnMatcherQuery : XlsxColumnMatcherQuery, IAsyncRequest<DataForMatcherUi>
    {
    }
}