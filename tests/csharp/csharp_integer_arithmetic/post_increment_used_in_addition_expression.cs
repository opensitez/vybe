// vybe-test: csharp/csharp_integer_arithmetic/post_increment_used_in_addition_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 4; __Check((value++ + 2).ToString(), "6");
