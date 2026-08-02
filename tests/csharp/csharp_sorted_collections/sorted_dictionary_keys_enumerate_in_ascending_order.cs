// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_keys_enumerate_in_ascending_order
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [3] = "c", [1] = "a", [2] = "b" }; foreach (var k in sd.Keys) Console.WriteLine(k);
