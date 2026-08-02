// vybe-test: csharp/csharp_regex_capture_groups/regex_match_value_returns_first_captured_group
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var match = System.Text.RegularExpressions.Regex.Match("id=42", @"id=(\d+)");
__Check((match.Groups[1].Value).ToString(), "42");
