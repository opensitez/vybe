// vybe-test: csharp/csharp_integer_arithmetic/chained_pre_increments_accumulate
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int left = 1; int right = 2; __Check((++left + ++right).ToString(), "5");
