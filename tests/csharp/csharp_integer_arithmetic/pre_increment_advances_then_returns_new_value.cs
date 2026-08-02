// vybe-test: csharp/csharp_integer_arithmetic/pre_increment_advances_then_returns_new_value
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 5; __Check((++value).ToString(), "6"); __Check((value).ToString(), "6");
