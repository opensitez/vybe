// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_negative_keys_sort_numerically
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [-1] = "neg", [0] = "zero", [1] = "pos" }; int first = 0; foreach (var k in sd.Keys) { first = k; break; } Console.WriteLine(first);
