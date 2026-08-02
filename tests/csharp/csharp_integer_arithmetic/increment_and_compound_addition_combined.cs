// vybe-test: csharp/csharp_integer_arithmetic/increment_and_compound_addition_combined
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 1; value++; value += 4; __Check((value).ToString(), "6");
