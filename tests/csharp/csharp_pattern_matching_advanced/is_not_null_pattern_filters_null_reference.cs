// vybe-test: csharp/csharp_pattern_matching_advanced/is_not_null_pattern_filters_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text = null; __Check((text is not null).ToString(), "False");
