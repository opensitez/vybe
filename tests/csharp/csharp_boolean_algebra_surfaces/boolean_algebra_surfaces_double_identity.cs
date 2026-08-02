// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
double seed = 11; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
