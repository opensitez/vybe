// vybe-test: csharp/csharp_integer_arithmetic/left_associativity_of_subtraction_without_parentheses
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10 - 4 - 2).ToString(), "4");
