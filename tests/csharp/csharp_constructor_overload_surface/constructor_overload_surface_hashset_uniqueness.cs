// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(67); set.Add(67); __Check((set.Count == 1).ToString(), "True");
