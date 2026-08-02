// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(63); set.Add(63); __Check((set.Count == 1).ToString(), "True");
