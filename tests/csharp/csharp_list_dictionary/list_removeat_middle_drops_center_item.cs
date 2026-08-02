// vybe-test: csharp/csharp_list_dictionary/list_removeat_middle_drops_center_item
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<int> { 10, 20, 30 }; list.RemoveAt(1); foreach (var x in list) Console.WriteLine(x);
