// vybe-test: csharp/csharp_integer_arithmetic/addition_and_subtraction_share_precedence_left_to_right
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10 + 5 - 3).ToString(), "12");
