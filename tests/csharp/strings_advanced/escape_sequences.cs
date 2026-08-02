// vybe-test: csharp/strings_advanced/escape_sequences
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("tab:\there").ToString(), "tab:\there");
__Check(("newline done").ToString(), "newline done");
