// vybe-test: csharp/strings_advanced/string_insert_remove
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "Hello World";
__Check((s.Insert(5, " Beautiful")).ToString(), "Hello Beautiful World");
__Check((s.Remove(5)).ToString(), "Hello");
__Check((s.Remove(5, 1)).ToString(), "HelloWorld");
