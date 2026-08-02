// vybe-test: csharp/collections_advanced/dict_iterate_keys_values
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 } };
foreach (var key in dict.Keys) Console.WriteLine(key);
