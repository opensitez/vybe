// vybe-test: csharp/csharp_numeric_types/short_range_min_max_values
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((short.MinValue).ToString(), "-32768"); __Check((short.MaxValue).ToString(), "32767");
