// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
int seed = 118; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
