// vybe-test: csharp/csharp_string_methods/ends_with_checks_suffix_match
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("body_suffix".EndsWith("suffix")).ToString(), "True");
