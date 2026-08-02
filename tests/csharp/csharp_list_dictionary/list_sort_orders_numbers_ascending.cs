// vybe-test: csharp/csharp_list_dictionary/list_sort_orders_numbers_ascending
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<int> { 3, 1, 2 }; list.Sort(); foreach (var x in list) Console.WriteLine(x);
