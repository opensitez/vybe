// vybe-test: csharp/csharp_integer_arithmetic/post_increment_returns_original_then_advances
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 5; __Check((value++).ToString(), "5"); __Check((value).ToString(), "6");
