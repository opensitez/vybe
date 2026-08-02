// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
int seed = 67; int right = seed + 1; __Check((seed < right).ToString(), "True");
