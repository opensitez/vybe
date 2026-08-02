// vybe-test: csharp/strings_advanced/string_join
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] words = { "hello", "world", "test" };
__Check((string.Join(", ", words)).ToString(), "hello, world, test");
