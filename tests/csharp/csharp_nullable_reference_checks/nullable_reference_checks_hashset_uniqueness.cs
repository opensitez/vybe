// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
var set = new System.Collections.Generic.HashSet<int>(); set.Add(58); set.Add(58); __Check((set.Count == 1).ToString(), "True");
