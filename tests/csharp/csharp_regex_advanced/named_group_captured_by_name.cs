// vybe-test: csharp/csharp_regex_advanced/named_group_captured_by_name
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var m = System.Text.RegularExpressions.Regex.Match("date=2024-06-15", @"(?<year>\d{4})-(?<month>\d{2})");
__Check((m.Groups["year"].Value).ToString(), "2024");
__Check((m.Groups["month"].Value).ToString(), "06");
