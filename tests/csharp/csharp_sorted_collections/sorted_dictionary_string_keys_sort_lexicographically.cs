// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_string_keys_sort_lexicographically
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using System.Collections.Generic; var sd = new SortedDictionary<string, int> { ["zebra"] = 1, ["apple"] = 2, ["mango"] = 3 }; foreach (var k in sd.Keys) Console.WriteLine(k);
