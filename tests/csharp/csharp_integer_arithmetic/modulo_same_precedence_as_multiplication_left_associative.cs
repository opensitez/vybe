// vybe-test: csharp/csharp_integer_arithmetic/modulo_same_precedence_as_multiplication_left_associative
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((20 % 6 % 2).ToString(), "0");
