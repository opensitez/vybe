// vybe-test: csharp/csharp_strings_ext/string_join
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts = new[] { "a", "b", "c" };
__Check((string.Join(", ", parts)).ToString(), "a, b, c");
