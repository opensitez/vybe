// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_dictionary_values_to_list_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

System.Collections.Generic.Dictionary<string, int> map = new() { ["a"] = 1, ["b"] = 2 };
System.Collections.Generic.List<int> values = new();
foreach (var kv in map) values.Add(kv.Value);
Console.WriteLine(values[0] + values[1]);
