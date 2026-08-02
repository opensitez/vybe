// vybe-test: csharp/csharp_integer_arithmetic/mixed_expression_with_post_increment_and_addition
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 3; __Check((value++ + value).ToString(), "7");
