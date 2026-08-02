// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_has_value_when_assigned
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal? d=1.5m; __Check((d.HasValue).ToString(), "True");
