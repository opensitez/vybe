// vybe-test: csharp/csharp_list_dictionary/list_foreach_emits_each_string_line
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<string> { "x", "y" }; foreach (var s in list) Console.WriteLine(s);
