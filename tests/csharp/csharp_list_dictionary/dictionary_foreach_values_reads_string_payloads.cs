// vybe-test: csharp/csharp_list_dictionary/dictionary_foreach_values_reads_string_payloads
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var map = new Dictionary<int, string> { [1] = "a", [2] = "b" }; foreach (var val in map.Values) Console.WriteLine(val);
