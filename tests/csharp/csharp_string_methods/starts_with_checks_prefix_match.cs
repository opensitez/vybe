// vybe-test: csharp/csharp_string_methods/starts_with_checks_prefix_match
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("prefix_body".StartsWith("prefix")).ToString(), "True");
