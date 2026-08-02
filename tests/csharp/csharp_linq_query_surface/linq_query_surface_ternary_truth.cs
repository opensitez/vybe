// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
int seed = 117; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
