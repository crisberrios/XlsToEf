﻿using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace XlsToEf.Import
{
    [Obsolete("ImportMatchingData obsolete, please use DataMatchesForImport", true)]
    public class ImportMatchingData : DataMatchesForImport
    {
        
    }

    public class DataMatchesForImport
    {
        public string FileName { get; set; }
        public string Sheet { get; set; }
        public List<XlsToEfColumnPair> Selected { get; set; }
        public string FilePath { get; set; }
    }

    public class XlsToEfColumnPair
    {
        public string XlsName { get; set; }

        public string EfName { get; set; }

        public XlsToEfColumnPair() 
        {
        }

        private XlsToEfColumnPair(string efName, string xlsName)
        {
            XlsName = xlsName;
            EfName = efName;
        }

        public static XlsToEfColumnPair Create(string efName, string xlsName)
        {
            return new XlsToEfColumnPair(efName, xlsName);
        }

        public static XlsToEfColumnPair Create<T>(Expression<Func<T>> propertyLambda, string xlsName)
        {
            var me = propertyLambda.Body as MemberExpression;

            if (me == null)
            {
                throw new ArgumentException("You must pass a lambda of the form: '() => Class.Property' or '() => object.Property'");
            }

            var name = me.Member.Name;
            return new XlsToEfColumnPair(name, xlsName);
        }
    }
}