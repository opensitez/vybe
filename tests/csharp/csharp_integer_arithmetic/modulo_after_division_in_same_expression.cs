// vybe-test: csharp/csharp_integer_arithmetic/modulo_after_division_in_same_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((17 / 5 % 3).ToString(), "0");
