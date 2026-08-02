// vybe-test: csharp/csharp_regex_advanced/matches_returns_all_non_overlapping_occurrences
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var matches = System.Text.RegularExpressions.Regex.Matches("a1 b2 c3", @"\d");
__Check((matches.Count).ToString(), "3");
