// vybe-test: csharp/csharp_regex_advanced/quantifier_plus_requires_one_or_more_digits
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Text.RegularExpressions.Regex.IsMatch("007", @"^\d+$")).ToString(), "True");
