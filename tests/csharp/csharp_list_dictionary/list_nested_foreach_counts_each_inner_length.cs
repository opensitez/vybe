// vybe-test: csharp/csharp_list_dictionary/list_nested_foreach_counts_each_inner_length
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var outer = new List<List<int>> { new List<int> { 1, 2 }, new List<int> { 3 } }; foreach (var inner in outer) Console.WriteLine(inner.Count);
