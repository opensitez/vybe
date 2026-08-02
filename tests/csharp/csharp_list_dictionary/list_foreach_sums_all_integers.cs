// vybe-test: csharp/csharp_list_dictionary/list_foreach_sums_all_integers
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using System.Collections.Generic; var list = new List<int> { 1, 2, 3 }; int sum = 0; foreach (var x in list) sum += x; Console.WriteLine(sum);
