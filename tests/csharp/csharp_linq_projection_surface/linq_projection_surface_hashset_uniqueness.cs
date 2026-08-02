// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(118); set.Add(118); __Check((set.Count == 1).ToString(), "True");
