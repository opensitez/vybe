// vybe-test: csharp/csharp_regex_advanced/replace_with_match_evaluator_transforms_each_match
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = System.Text.RegularExpressions.Regex.Replace(
    "a1b2c3", @"\d",
    m => ((int.Parse(m.Value)*2)).ToString());
__Check((result).ToString(), "a2b4c6");
