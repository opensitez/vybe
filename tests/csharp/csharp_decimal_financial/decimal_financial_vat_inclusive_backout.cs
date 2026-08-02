// vybe-test: csharp/csharp_decimal_financial/decimal_financial_vat_inclusive_backout
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal gross=119.00m; decimal vatRate=0.19m; __Check((gross/(1m+vatRate)).ToString(), "100");
