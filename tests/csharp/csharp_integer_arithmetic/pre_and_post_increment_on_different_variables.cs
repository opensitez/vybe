// vybe-test: csharp/csharp_integer_arithmetic/pre_and_post_increment_on_different_variables
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int left = 2; int right = 5; __Check((++left + right++).ToString(), "8"); __Check((left).ToString(), "3"); __Check((right).ToString(), "6");
