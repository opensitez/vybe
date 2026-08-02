// vybe-test: csharp/strings_advanced/string_isnullorempty
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.IsNullOrEmpty("")).ToString(), "True");
__Check((string.IsNullOrEmpty(null)).ToString(), "True");
__Check((string.IsNullOrEmpty("hello")).ToString(), "False");
