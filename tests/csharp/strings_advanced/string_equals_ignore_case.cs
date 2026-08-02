// vybe-test: csharp/strings_advanced/string_equals_ignore_case
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Equals("Hello", "hello", StringComparison.OrdinalIgnoreCase)).ToString(), "True");
__Check((string.Equals("Hello", "hello")).ToString(), "False");
