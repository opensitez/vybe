// vybe-test: csharp/csharp_map_set_collections/hashset_rejects_duplicate_value_addition
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var set = new HashSet<int>(); __Check((set.Add(3)).ToString(), "True"); __Check((set.Add(3)).ToString(), "False");
