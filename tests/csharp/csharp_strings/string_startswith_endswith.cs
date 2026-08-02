// vybe-test: csharp/csharp_strings/string_startswith_endswith
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello".StartsWith("hel")).ToString(), "True");
__Check(("hello".EndsWith("llo")).ToString(), "True");
__Check(("hello".StartsWith("xyz")).ToString(), "False");
