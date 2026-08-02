// vybe-test: csharp/csharp_string_methods/contains_returns_true_for_present_substring
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("foobar".Contains("oba")).ToString(), "True");
