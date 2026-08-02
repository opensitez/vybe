// vybe-test: csharp/csharp_map_set_collections/dictionary_keys_collection_can_be_enumerated
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; foreach (var key in map.Keys) Console.WriteLine(key);
