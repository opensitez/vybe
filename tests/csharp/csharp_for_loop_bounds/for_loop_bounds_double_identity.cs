// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
double seed = 45; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
