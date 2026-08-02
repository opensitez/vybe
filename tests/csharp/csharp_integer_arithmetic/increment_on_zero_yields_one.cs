// vybe-test: csharp/csharp_integer_arithmetic/increment_on_zero_yields_one
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 0; value++; __Check((value).ToString(), "1");
