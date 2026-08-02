// vybe-test: csharp/csharp_if_else_branching/if_else_branching_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// if_else_branching
var set = new System.Collections.Generic.HashSet<int>(); set.Add(44); set.Add(44); __Check((set.Count == 1).ToString(), "True");
