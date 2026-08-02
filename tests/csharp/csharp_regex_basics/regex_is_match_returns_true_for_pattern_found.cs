// vybe-test: csharp/csharp_regex_basics/regex_is_match_returns_true_for_pattern_found
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Text.RegularExpressions.Regex.IsMatch("hello123","[0-9]+")).ToString(), "True");
