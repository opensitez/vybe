// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
int seed = 52; __Check((seed + 1 > seed).ToString(), "True");
