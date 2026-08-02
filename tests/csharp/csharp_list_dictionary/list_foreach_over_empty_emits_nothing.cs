// vybe-test: csharp/csharp_list_dictionary/list_foreach_over_empty_emits_nothing
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<int>(); foreach (var x in list) Console.WriteLine(x); Console.WriteLine("done");
