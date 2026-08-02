// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
int seed = 45; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
