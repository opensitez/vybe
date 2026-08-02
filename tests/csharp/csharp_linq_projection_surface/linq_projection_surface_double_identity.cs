// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
double seed = 118; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
