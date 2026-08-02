// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
var set = new System.Collections.Generic.HashSet<int>(); set.Add(29); set.Add(29); __Check((set.Count == 1).ToString(), "True");
