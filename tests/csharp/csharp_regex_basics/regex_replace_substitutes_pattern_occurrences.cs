// vybe-test: csharp/csharp_regex_basics/regex_replace_substitutes_pattern_occurrences
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r=System.Text.RegularExpressions.Regex.Replace("a1b2c3","[0-9]","#");
__Check((r).ToString(), "a#b#c#");
