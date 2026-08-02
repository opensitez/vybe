// vybe-test: csharp/csharp_regex_capture_groups/regex_replace_substitutes_all_occurrences_with_replacement_text
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var text = System.Text.RegularExpressions.Regex.Replace("a-b-c", "-", "_");
__Check((text).ToString(), "a_b_c");
