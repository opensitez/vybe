// vybe-test: csharp/csharp_integer_arithmetic/addition_of_two_positive_literals
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((12 + 8).ToString(), "20");
