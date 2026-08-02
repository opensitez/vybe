// vybe-test: csharp/csharp_decimal_semantics/decimal_equality_compares_numeric_value_not_reference
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal left = 1.0m;
decimal right = 1.00m;
__Check((left == right).ToString(), "True");
