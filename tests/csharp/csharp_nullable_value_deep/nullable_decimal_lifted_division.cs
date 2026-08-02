// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_lifted_division
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal? a=7.5m; decimal? b=2.5m; __Check((a/b).ToString(), "3");
