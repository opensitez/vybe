// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(49); set.Add(49); __Check((set.Count == 1).ToString(), "True");
