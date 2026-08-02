// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// floating_point_literals_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(16); set.Add(16); __Check((set.Count == 1).ToString(), "True");
