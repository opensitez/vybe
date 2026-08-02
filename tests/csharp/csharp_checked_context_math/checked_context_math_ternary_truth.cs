// vybe-test: csharp/csharp_checked_context_math/checked_context_math_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
int seed = 12; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
