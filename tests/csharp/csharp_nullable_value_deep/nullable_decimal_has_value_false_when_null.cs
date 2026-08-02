// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_has_value_false_when_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal? d=null; __Check((d.HasValue).ToString(), "False");
