// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
double seed = 30; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
