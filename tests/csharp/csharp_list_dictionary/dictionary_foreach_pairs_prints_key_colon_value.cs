// vybe-test: csharp/csharp_list_dictionary/dictionary_foreach_pairs_prints_key_colon_value
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var map = new Dictionary<string, int> { ["b"] = 2, ["a"] = 1 }; foreach (var pair in map) Console.WriteLine(pair.Key + ":" + pair.Value);
