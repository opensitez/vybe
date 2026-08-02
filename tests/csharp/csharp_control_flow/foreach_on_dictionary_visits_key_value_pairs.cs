// vybe-test: csharp/csharp_control_flow/foreach_on_dictionary_visits_key_value_pairs
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

var map = new System.Collections.Generic.Dictionary<string, int> { ["x"] = 1 };
int total = 0;
foreach (var pair in map) total += pair.Value;
Console.WriteLine(total);
