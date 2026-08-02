// vybe-test: csharp/strings_advanced/string_contains_startswith_endswith
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "Hello World";
__Check((s.Contains("lo Wo")).ToString(), "True");
__Check((s.StartsWith("Hello")).ToString(), "True");
__Check((s.EndsWith("World")).ToString(), "True");
__Check((s.StartsWith("World")).ToString(), "False");
