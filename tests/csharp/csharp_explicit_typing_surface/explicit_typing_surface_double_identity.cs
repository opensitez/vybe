// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// explicit_typing_surface
double seed = 60; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
