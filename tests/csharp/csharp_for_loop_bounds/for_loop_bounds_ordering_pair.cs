// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
int seed = 45; int right = seed + 1; __Check((seed < right).ToString(), "True");
