// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
var set = new System.Collections.Generic.HashSet<int>(); set.Add(53); set.Add(53); __Check((set.Count == 1).ToString(), "True");
