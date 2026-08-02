// vybe-test: csharp/csharp_integer_arithmetic/division_of_equal_operands_yields_one
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((15 / 15).ToString(), "1");
