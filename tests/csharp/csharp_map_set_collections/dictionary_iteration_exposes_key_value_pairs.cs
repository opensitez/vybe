// vybe-test: csharp/csharp_map_set_collections/dictionary_iteration_exposes_key_value_pairs
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using System.Collections.Generic; var map = new Dictionary<string, int> { ["b"] = 2, ["a"] = 1 }; foreach (var pair in map) Console.WriteLine(pair.Key + ":" + pair.Value);
