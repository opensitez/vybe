// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
var set = new System.Collections.Generic.HashSet<int>(); set.Add(115); set.Add(115); __Check((set.Count == 1).ToString(), "True");
