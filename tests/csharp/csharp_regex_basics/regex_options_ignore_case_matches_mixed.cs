// vybe-test: csharp/csharp_regex_basics/regex_options_ignore_case_matches_mixed
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool r=System.Text.RegularExpressions.Regex.IsMatch("Hello","hello",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
__Check((r).ToString(), "True");
