// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
var tuple = (left: 13, right: 14); __Check((tuple.left < tuple.right).ToString(), "True");
