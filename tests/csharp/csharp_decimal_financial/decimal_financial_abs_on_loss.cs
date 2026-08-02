// vybe-test: csharp/csharp_decimal_financial/decimal_financial_abs_on_loss
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal pnl=-125.40m; __Check((System.Math.Abs(pnl)).ToString(), "125.40");
