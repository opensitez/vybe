// vybe-test: csharp/csharp_integer_arithmetic/subtracting_negative_equivalent_to_addition
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((7 - (-3)).ToString(), "10");
