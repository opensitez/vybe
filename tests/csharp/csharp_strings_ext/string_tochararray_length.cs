// vybe-test: csharp/csharp_strings_ext/string_tochararray_length
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "hello";
__Check((s.Length).ToString(), "5");
__Check((s[0]).ToString(), "h");
__Check((s[4]).ToString(), "o");
