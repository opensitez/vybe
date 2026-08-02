// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// floating_point_literals_surface
int seed = 16; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
