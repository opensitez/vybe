// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_lifted_addition_one_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal? a=1m; decimal? b=null; __Check(((a+b).HasValue).ToString(), "False");
