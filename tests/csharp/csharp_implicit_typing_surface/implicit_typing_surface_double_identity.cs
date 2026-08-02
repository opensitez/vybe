// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// implicit_typing_surface
double seed = 59; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
