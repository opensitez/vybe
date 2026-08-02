// vybe-test: csharp/csharp_decimal_financial/decimal_financial_unary_negate_credit
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal credit=250.75m; __Check((-credit).ToString(), "-250.75");
