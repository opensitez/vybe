// vybe-test: csharp/csharp_regex_advanced/character_class_matches_any_listed_char
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var m = System.Text.RegularExpressions.Regex.Match("hello", @"[aeiou]");
__Check((m.Value).ToString(), "e");
