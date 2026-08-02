// vybe-test: csharp/csharp_regex_capture_groups/regex_options_ignore_case_matches_differing_casing
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool ok = System.Text.RegularExpressions.Regex.IsMatch(
    "Hello",
    "hello",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
__Check((ok).ToString(), "True");
