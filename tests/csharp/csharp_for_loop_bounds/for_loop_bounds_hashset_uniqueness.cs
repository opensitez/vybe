// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
var set = new System.Collections.Generic.HashSet<int>(); set.Add(45); set.Add(45); __Check((set.Count == 1).ToString(), "True");
