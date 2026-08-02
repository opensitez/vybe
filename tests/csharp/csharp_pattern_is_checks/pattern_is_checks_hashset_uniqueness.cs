// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
var set = new System.Collections.Generic.HashSet<int>(); set.Add(41); set.Add(41); __Check((set.Count == 1).ToString(), "True");
