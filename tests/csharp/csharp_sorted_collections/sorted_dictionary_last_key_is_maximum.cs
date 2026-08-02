// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_last_key_is_maximum
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using System.Collections.Generic; var sd = new SortedDictionary<int, string> { [10] = "ten", [2] = "two", [7] = "seven" }; int last = 0; foreach (var k in sd.Keys) last = k; Console.WriteLine(last);
