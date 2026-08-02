// vybe-test: csharp/csharp_integer_arithmetic/addition_chain_of_three_terms_is_left_associative
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((1 + 2 + 3).ToString(), "6");
