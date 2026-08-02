// vybe-test: csharp/csharp_dictionary_operations/keys_collection_enumerates_all_inserted_keys
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

var d = new System.Collections.Generic.Dictionary<string,int>{{"x",1},{"y",2}};
var keys = new System.Collections.Generic.List<string>(d.Keys);
keys.Sort();
foreach(var k in keys) Console.WriteLine(k);
