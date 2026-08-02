// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_foreach_yields_sorted_pairs
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using System.Collections.Generic; var sd = new SortedDictionary<string, int> { ["b"] = 2, ["a"] = 1, ["c"] = 3 }; foreach (var p in sd) Console.WriteLine(p.Key + ":" + p.Value);
