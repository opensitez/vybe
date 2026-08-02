// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(30); set.Add(30); __Check((set.Count == 1).ToString(), "True");
