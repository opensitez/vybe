// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
int seed = 13; __Check((seed + 1 > seed).ToString(), "True");
