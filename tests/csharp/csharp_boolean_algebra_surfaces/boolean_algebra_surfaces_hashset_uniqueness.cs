// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
var set = new System.Collections.Generic.HashSet<int>(); set.Add(11); set.Add(11); __Check((set.Count == 1).ToString(), "True");
