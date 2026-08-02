// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
double seed = 67; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
