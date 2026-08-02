// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
var tuple = (left: 22, right: 23); __Check((tuple.left < tuple.right).ToString(), "True");
