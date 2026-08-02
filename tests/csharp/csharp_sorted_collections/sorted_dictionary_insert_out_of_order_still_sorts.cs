// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_insert_out_of_order_still_sorts
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using System.Collections.Generic; var sd = new SortedDictionary<int, int>(); sd[30] = 3; sd[10] = 1; sd[20] = 2; int sum = 0; foreach (var p in sd) sum += p.Key; Console.WriteLine(sum);
