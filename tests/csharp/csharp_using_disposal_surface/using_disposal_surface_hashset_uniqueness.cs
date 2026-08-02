// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(52); set.Add(52); __Check((set.Count == 1).ToString(), "True");
