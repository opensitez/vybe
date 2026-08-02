// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
int seed = 30; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
