// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
var values = new System.Collections.Generic.List<int> { 13, 14, 13 }; __Check((values.Count == 3).ToString(), "True");
