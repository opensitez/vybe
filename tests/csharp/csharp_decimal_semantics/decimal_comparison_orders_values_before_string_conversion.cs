// vybe-test: csharp/csharp_decimal_semantics/decimal_comparison_orders_values_before_string_conversion
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal low = 1.2m;
decimal high = 1.3m;
__Check((low < high).ToString(), "True");
__Check((high > low).ToString(), "True");
