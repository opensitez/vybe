// vybe-test: csharp/csharp_strings/string_comparison
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("abc" == "abc").ToString(), "True");
__Check(("abc" != "xyz").ToString(), "True");
