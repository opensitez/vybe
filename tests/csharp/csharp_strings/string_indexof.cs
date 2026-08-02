// vybe-test: csharp/csharp_strings/string_indexof
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello world".IndexOf("world")).ToString(), "6");
__Check(("hello world".IndexOf("xyz")).ToString(), "-1");
