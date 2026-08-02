// vybe-test: csharp/csharp_list_dictionary/list_insert_at_zero_prefixes_new_item
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<int> { 2, 3 }; list.Insert(0, 1); foreach (var x in list) Console.WriteLine(x);
