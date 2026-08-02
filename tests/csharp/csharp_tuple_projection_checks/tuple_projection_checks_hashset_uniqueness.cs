// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
var set = new System.Collections.Generic.HashSet<int>(); set.Add(36); set.Add(36); __Check((set.Count == 1).ToString(), "True");
