// vybe-test: csharp/csharp_bcl_collections/sorted_dictionary_enumerator_yields_keys_in_ascending_order
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

var map = new System.Collections.Generic.SortedDictionary<int, string>();
map[3] = "c";
map[1] = "a";
int firstKey = 0;
foreach (var pair in map) { firstKey = pair.Key; break; }
Console.WriteLine(firstKey);
