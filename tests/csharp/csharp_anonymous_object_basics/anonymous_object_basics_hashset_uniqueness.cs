// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
var set = new System.Collections.Generic.HashSet<int>(); set.Add(38); set.Add(38); __Check((set.Count == 1).ToString(), "True");
