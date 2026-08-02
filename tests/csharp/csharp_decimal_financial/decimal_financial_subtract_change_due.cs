// vybe-test: csharp/csharp_decimal_financial/decimal_financial_subtract_change_due
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal paid=50.00m; decimal total=37.42m; __Check((paid-total).ToString(), "12.58");
