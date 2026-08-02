// vybe-test: csharp/csharp_list_dictionary/dictionary_foreach_keys_yields_inserted_string_keys
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }; foreach (var key in map.Keys) Console.WriteLine(key);
