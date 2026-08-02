// vybe-test: csharp/strings_advanced/string_replace
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "hello world";
__Check((s.Replace("world", "there")).ToString(), "hello there");
__Check((s.Replace("l", "L")).ToString(), "heLLo worLd");
