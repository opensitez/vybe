// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// serialization_json_surface
int seed = 91; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
