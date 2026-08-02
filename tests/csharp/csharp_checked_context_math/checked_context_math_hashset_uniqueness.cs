// vybe-test: csharp/csharp_checked_context_math/checked_context_math_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
var set = new System.Collections.Generic.HashSet<int>(); set.Add(12); set.Add(12); __Check((set.Count == 1).ToString(), "True");
