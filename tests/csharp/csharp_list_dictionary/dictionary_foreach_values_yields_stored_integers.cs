// vybe-test: csharp/csharp_list_dictionary/dictionary_foreach_values_yields_stored_integers
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var map = new Dictionary<string, int> { ["a"] = 10, ["b"] = 20 }; foreach (var val in map.Values) Console.WriteLine(val);
