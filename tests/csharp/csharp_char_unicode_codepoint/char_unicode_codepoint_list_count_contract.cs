// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
var values = new System.Collections.Generic.List<int> { 22, 23, 22 }; __Check((values.Count == 3).ToString(), "True");
