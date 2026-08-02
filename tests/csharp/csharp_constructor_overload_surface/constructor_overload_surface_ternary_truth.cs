// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
int seed = 67; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
