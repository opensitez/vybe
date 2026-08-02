// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// explicit_typing_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(60); set.Add(60); __Check((set.Count == 1).ToString(), "True");
