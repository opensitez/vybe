// vybe-test: csharp/csharp_nullable_value_deep/nullable_decimal_value_reads_fraction
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal? d=2.25m; __Check((d.Value).ToString(), "2.25");
