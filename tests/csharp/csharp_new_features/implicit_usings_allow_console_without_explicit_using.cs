// vybe-test: csharp/csharp_new_features/implicit_usings_allow_console_without_explicit_using
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// new_features
__Check((42).ToString(), "42");
