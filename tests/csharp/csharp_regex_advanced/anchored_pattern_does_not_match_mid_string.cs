// vybe-test: csharp/csharp_regex_advanced/anchored_pattern_does_not_match_mid_string
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Text.RegularExpressions.Regex.IsMatch("abc", @"^\d+$")).ToString(), "False");
