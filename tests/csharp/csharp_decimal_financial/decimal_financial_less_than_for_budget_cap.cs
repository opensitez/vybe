// vybe-test: csharp/csharp_decimal_financial/decimal_financial_less_than_for_budget_cap
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal spent=999.99m; decimal cap=1000.00m; __Check((spent<cap).ToString(), "True");
