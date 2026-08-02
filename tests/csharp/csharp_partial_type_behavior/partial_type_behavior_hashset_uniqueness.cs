// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
var set = new System.Collections.Generic.HashSet<int>(); set.Add(70); set.Add(70); __Check((set.Count == 1).ToString(), "True");
