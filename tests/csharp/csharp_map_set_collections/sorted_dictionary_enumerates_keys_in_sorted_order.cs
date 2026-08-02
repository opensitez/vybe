// vybe-test: csharp/csharp_map_set_collections/sorted_dictionary_enumerates_keys_in_sorted_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using System.Collections.Generic; var map = new SortedDictionary<string, int> { ["b"] = 2, ["a"] = 1 }; foreach (var pair in map) Console.WriteLine(pair.Key + ":" + pair.Value);
