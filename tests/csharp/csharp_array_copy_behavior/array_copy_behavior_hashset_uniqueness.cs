// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
var set = new System.Collections.Generic.HashSet<int>(); set.Add(26); set.Add(26); __Check((set.Count == 1).ToString(), "True");
