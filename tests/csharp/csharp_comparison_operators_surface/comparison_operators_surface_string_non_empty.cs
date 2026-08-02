// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
string feature = "comparison_operators_surface"; __Check((feature.Length > 0).ToString(), "True");
