// vybe-test: csharp/csharp_decimal_financial/decimal_financial_penny_allocation_first
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal total=0.10m; int parts=3; decimal share=decimal.Truncate(total/parts*100m)/100m; __Check((share).ToString(), "0.03");
