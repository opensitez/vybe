// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
var set = new System.Collections.Generic.HashSet<int>(); set.Add(22); set.Add(22); __Check((set.Count == 1).ToString(), "True");
