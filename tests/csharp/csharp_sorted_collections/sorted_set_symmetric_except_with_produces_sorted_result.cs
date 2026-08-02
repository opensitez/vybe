// vybe-test: csharp/csharp_sorted_collections/sorted_set_symmetric_except_with_produces_sorted_result
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var ss = new SortedSet<int> { 1, 2, 3 }; ss.SymmetricExceptWith(new[] { 2, 3, 4 }); __Check((ss.Contains(1)).ToString(), "True"); __Check((ss.Contains(4)).ToString(), "True"); __Check((ss.Contains(2)).ToString(), "False");
