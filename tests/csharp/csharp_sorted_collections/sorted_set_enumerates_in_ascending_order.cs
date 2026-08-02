// vybe-test: csharp/csharp_sorted_collections/sorted_set_enumerates_in_ascending_order
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using System.Collections.Generic; var ss = new SortedSet<int> { 5, 1, 3, 4, 2 }; foreach (var x in ss) Console.WriteLine(x);
