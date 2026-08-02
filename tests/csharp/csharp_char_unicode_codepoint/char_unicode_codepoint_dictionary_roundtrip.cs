// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
var map = new System.Collections.Generic.Dictionary<int, int>(); map[22] = 23; __Check((map.ContainsKey(22) && map[22] == 23).ToString(), "True");
