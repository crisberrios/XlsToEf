using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;
using Shouldly;
using XlsToEf.Import;
using XlsToEf.Tests.ImportHelperFiles;
using XlsToEf.Tests.Models;

namespace XlsToEf.Tests
{
    public class TestPropertyOverriderDbTests : DbTestBase
    {
        public async Task ShouldPopulateDestinationFromSource()
        {
            var dbContext = GetDb();
            var originalCategory = new ProductCategory {CategoryCode = "abc"};
            var newCategory = new ProductCategory {CategoryCode = "def"};
            const string awesomeNewName = "Awesome New Name";

            var destination = new Product {ProductCategory = originalCategory};

            PersistToDatabase(originalCategory, newCategory);

            var matches = new Dictionary<string, string>
            {
                {"ProductCategory", "cat"},
                {"ProductName", "name"},

            };


            var excelRow = new Dictionary<string, string>
            {
                {"cat", newCategory.CategoryCode},
                {"name", awesomeNewName},

            };

            var overrider = new ProductPropertyOverrider<Product>(dbContext);
            await overrider.UpdateProperties(destination, matches, excelRow);
            destination.ProductCategory.CategoryCode.ShouldBe(newCategory.CategoryCode);
            destination.ProductName.ShouldBe(awesomeNewName);
        }

        public async Task ShouldImportWithOverrider()
        {
            var dbContext = GetDb();
            var categoryName = "Cookies";
            var categoryToSelect = new ProductCategory { CategoryCode = "CK", CategoryName = categoryName };
            var unselectedCategory = new ProductCategory { CategoryCode = "UC", CategoryName = "Unrelated Category" };

            PersistToDatabase(categoryToSelect, unselectedCategory);

            var existingProduct = new Product {ProductCategoryId = unselectedCategory.Id, ProductName = "Vanilla Wafers"};
            PersistToDatabase(existingProduct);


            var overrider = new ProductPropertyOverrider<Product>(dbContext);


            var excelIoWrapper = new FakeExcelIo();
            excelIoWrapper.Rows.Clear();
            var cookieType = "Mint Cookies";
            excelIoWrapper.Rows.Add(new Dictionary<string, string>
            {
                {"xlsCol1", "CK"},
                {"xlsCol2", cookieType },
                {"xlsCol5", "" },

            });
            var updatedCookieName = "Strawberry Wafers";
            excelIoWrapper.Rows.Add(new Dictionary<string, string>
            {
                {"xlsCol1", "CK"},
                {"xlsCol2", updatedCookieName },
                {"xlsCol5", existingProduct.Id.ToString() },

            });


            var importer = new XlsxToTableImporter(dbContext, excelIoWrapper);

            var importMatchingData = new ImportMatchingOrderData
            {
                FileName = "foo.xlsx",
                Sheet = "mysheet",
                Selected =
                    new Dictionary<string, string>
                    {
                        {"Id", "xlsCol5"},
                        {"ProductCategory", "xlsCol1"},
                        {"ProductName", "xlsCol2"},
                    }
            };


            Func<int, Expression<Func<Product, bool>>> finderExpression = selectorValue => entity => entity.Id.Equals( selectorValue);
            var result =  await importer.ImportColumnData(importMatchingData, finderExpression, overridingMapper: overrider, recordMode: RecordMode.Upsert);

            var newItem = GetDb().Set<Product>().Include(x => x.ProductCategory).First(x => x.ProductName == cookieType);
            newItem.ProductCategory.CategoryName.ShouldBe("Cookies");
            newItem.ProductName.ShouldBe(cookieType);

            var updated = GetDb().Set<Product>().Include(x => x.ProductCategory).First(x => x.Id == existingProduct.Id);
            updated.ProductName.ShouldBe(updatedCookieName);
        }


        private class ProductPropertyOverrider<T> : UpdatePropertyOverrider<T> where T : Product
        {
            private readonly DbContext _context;

            public ProductPropertyOverrider(DbContext context)
            {
                _context = context;
            }

            public override async Task UpdateProperties(T destination1, Dictionary<string, string> matches,
                Dictionary<string, string> excelRow)
            {
                {
                    var product = new Product();
                    var productCategoryPropertyName =
                        PropertyNameHelper.GetPropertyName(() => product.ProductCategory);
                    var productPropertyName = PropertyNameHelper.GetPropertyName(() => product.ProductName);

                    foreach (var destinationProperty in matches.Keys)
                    {
                        var xlsxColumnName = matches[destinationProperty];
                        var value = excelRow[xlsxColumnName];
                        if (destinationProperty == productCategoryPropertyName)
                        {
                            var newCategory =
                                await _context.Set<ProductCategory>().Where(x => x.CategoryCode == value).FirstAsync();
                            if (newCategory == null)
                                throw new RowParseException("Category Code does not match a category");
                            destination1.ProductCategory = newCategory;
                        }
                        else if (destinationProperty == productPropertyName)
                        {
                            destination1.ProductName = value;
                        }
                    }
                }
            }
        }
    }
}