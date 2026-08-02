// vybe-test: csharp/csharp_list_dictionary/list_sort_orders_words_alphabetically
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<string> { "c", "a", "b" }; list.Sort(); foreach (var s in list) Console.WriteLine(s);
