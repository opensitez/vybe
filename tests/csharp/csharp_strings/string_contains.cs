// vybe-test: csharp/csharp_strings/string_contains
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello world".Contains("world")).ToString(), "True");
__Check(("hello world".Contains("xyz")).ToString(), "False");
