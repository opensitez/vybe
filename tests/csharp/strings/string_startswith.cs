// vybe-test: csharp/strings/string_startswith
// origin: languages/csharp/tests/csharp/test_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello".StartsWith("hel")).ToString(), "True");
        __Check(("hello".StartsWith("xyz")).ToString(), "False");
