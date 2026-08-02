// vybe-test: csharp/csharp_pattern_matching_advanced/is_pattern_rejects_non_matching_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object item = 42; __Check((item is string text).ToString(), "False");
