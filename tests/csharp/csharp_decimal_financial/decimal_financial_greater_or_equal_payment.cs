// vybe-test: csharp/csharp_decimal_financial/decimal_financial_greater_or_equal_payment
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal due=50.00m; decimal paid=50.00m; __Check((paid>=due).ToString(), "True");
