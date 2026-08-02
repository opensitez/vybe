// vybe-test: csharp/csharp_map_set_collections/hashset_remove_erases_existing_member
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var set = new HashSet<int> { 1, 2 }; set.Remove(1); __Check((set.Contains(1)).ToString(), "False");
