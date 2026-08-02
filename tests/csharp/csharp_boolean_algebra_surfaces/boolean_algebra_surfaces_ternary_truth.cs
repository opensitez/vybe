// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
int seed = 11; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
