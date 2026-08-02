// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
var set = new System.Collections.Generic.HashSet<int>(); set.Add(69); set.Add(69); __Check((set.Count == 1).ToString(), "True");
