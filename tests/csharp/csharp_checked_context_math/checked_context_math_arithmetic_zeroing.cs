// vybe-test: csharp/csharp_checked_context_math/checked_context_math_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
int seed = 12; __Check((seed - seed == 0).ToString(), "True");
