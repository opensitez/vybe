// vybe-test: csharp/csharp_regex_advanced/multiline_option_applies_caret_to_each_line_start
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var matches = System.Text.RegularExpressions.Regex.Matches(
    "start\nnew line", @"^[a-z]",
    System.Text.RegularExpressions.RegexOptions.Multiline);
__Check((matches.Count).ToString(), "2");
