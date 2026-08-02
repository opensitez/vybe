// vybe-test: csharp/csharp_string_split_join/string_split_join_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
var set = new System.Collections.Generic.HashSet<int>(); set.Add(21); set.Add(21); __Check((set.Count == 1).ToString(), "True");
