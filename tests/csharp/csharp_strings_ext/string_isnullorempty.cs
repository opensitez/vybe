// vybe-test: csharp/csharp_strings_ext/string_isnullorempty
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.IsNullOrEmpty(null)).ToString(), "True");
__Check((string.IsNullOrEmpty("")).ToString(), "True");
__Check((string.IsNullOrEmpty("hello")).ToString(), "False");
