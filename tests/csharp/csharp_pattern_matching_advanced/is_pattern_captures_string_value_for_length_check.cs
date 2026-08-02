// vybe-test: csharp/csharp_pattern_matching_advanced/is_pattern_captures_string_value_for_length_check
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object item = "alpha"; if (item is string text) __Check((text.Length).ToString(), "5");
