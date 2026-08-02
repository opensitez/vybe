// vybe-test: csharp/csharp_list_dictionary/list_insert_at_middle_splits_sequence
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<string> { "a", "c" }; list.Insert(1, "b"); foreach (var s in list) Console.WriteLine(s);
