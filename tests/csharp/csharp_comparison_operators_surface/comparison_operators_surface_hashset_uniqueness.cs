// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(13); set.Add(13); __Check((set.Count == 1).ToString(), "True");
