// vybe-test: csharp/csharp_numeric_types/int_min_value_is_negative_two_billion_one_forty_eight_million
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((int.MinValue).ToString(), "-2147483648");
