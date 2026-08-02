// vybe-test: csharp/csharp_decimal_financial/decimal_financial_interest_simple_one_year
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal p=5000m; decimal r=0.05m; __Check((p+p*r).ToString(), "5250.00");
