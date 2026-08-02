// vybe-test: csharp/strings_advanced/string_indexof_lastindexof
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "abcabc";
__Check((s.IndexOf("bc")).ToString(), "1");
__Check((s.LastIndexOf("bc")).ToString(), "4");
__Check((s.IndexOf("xyz")).ToString(), "-1");
