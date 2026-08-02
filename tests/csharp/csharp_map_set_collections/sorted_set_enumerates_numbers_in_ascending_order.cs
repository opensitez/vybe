// vybe-test: csharp/csharp_map_set_collections/sorted_set_enumerates_numbers_in_ascending_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using System.Collections.Generic; var set = new SortedSet<int> { 5, 1, 3 }; foreach (var item in set) Console.WriteLine(item);
