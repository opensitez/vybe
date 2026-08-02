// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_null_coalescing
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal? d=null; __Check((d??3.5m).ToString(), "3.5");
