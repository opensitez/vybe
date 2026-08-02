// vybe-test: csharp/csharp_strings/string_substring
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello world".Substring(6)).ToString(), "world");
__Check(("hello world".Substring(0, 5)).ToString(), "hello");
