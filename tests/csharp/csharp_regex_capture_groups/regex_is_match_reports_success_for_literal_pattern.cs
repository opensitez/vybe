// vybe-test: csharp/csharp_regex_capture_groups/regex_is_match_reports_success_for_literal_pattern
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Text.RegularExpressions.Regex.IsMatch("abc123", @"\d+")).ToString(), "True");
