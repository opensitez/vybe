// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_comparison_less_than
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal? a=1.2m; decimal? b=1.3m; __Check((a<b).ToString(), "True");
