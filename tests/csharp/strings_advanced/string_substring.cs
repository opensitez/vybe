// vybe-test: csharp/strings_advanced/string_substring
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "Hello World";
__Check((s.Substring(6)).ToString(), "World");
__Check((s.Substring(0, 5)).ToString(), "Hello");
