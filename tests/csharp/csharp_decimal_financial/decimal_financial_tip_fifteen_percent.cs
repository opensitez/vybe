// vybe-test: csharp/csharp_decimal_financial/decimal_financial_tip_fifteen_percent
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal meal=47.80m; __Check((meal*0.15m).ToString(), "7.170");
