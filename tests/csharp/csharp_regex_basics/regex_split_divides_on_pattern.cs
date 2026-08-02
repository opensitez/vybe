// vybe-test: csharp/csharp_regex_basics/regex_split_divides_on_pattern
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts=System.Text.RegularExpressions.Regex.Split("one1two2three","[0-9]");
__Check((parts.Length).ToString(), "3"); __Check((parts[1]).ToString(), "two");
