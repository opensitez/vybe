// vybe-test: csharp/csharp_string_culture/to_lower_with_invariant_culture
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("HELLO".ToLowerInvariant()).ToString(), "hello");
