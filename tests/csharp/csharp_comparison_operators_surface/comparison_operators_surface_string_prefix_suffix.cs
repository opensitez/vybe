// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
string feature = "comparison_operators_surface"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
