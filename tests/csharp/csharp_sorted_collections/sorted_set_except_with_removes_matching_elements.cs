// vybe-test: csharp/csharp_sorted_collections/sorted_set_except_with_removes_matching_elements
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3, 4 }; ss.ExceptWith(new[] { 2, 4 }); __Check((ss.Count).ToString(), "2"); __Check((ss.Contains(1)).ToString(), "True");
