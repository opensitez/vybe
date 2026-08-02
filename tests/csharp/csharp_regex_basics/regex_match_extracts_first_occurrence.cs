// vybe-test: csharp/csharp_regex_basics/regex_match_extracts_first_occurrence
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var m=System.Text.RegularExpressions.Regex.Match("abc123def","[0-9]+");
__Check((m.Value).ToString(), "123");
