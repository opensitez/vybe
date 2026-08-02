// vybe-test: csharp/csharp_decimal_semantics/decimal_unary_minus_negates_value
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal balance = 12.5m; __Check((-balance).ToString(), "-12.5");
