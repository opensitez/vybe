// vybe-test: csharp/csharp_string_methods/to_upper_lower_changes_case_of_all_letters
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("Hello".ToUpper()).ToString(), "HELLO"); __Check(("Hello".ToLower()).ToString(), "hello");
