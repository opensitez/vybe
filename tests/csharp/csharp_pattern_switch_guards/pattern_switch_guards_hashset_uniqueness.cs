// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
var set = new System.Collections.Generic.HashSet<int>(); set.Add(42); set.Add(42); __Check((set.Count == 1).ToString(), "True");
