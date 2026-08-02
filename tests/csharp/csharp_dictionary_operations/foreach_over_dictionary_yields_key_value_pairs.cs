// vybe-test: csharp/csharp_dictionary_operations/foreach_over_dictionary_yields_key_value_pairs
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

var d = new System.Collections.Generic.Dictionary<int,int>{{1,10}};
foreach(var pair in d) Console.WriteLine(pair.Key + ":" + pair.Value);
