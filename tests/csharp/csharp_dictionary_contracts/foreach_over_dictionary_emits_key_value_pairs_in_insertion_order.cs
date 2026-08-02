// vybe-test: csharp/csharp_dictionary_contracts/foreach_over_dictionary_emits_key_value_pairs_in_insertion_order
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

using System.Collections.Generic;
var map = new Dictionary<string, int> {
    ["b"] = 2,
    ["a"] = 1,
    ["c"] = 3
};
foreach (var entry in map) {
    Console.WriteLine(entry.Key + ":" + entry.Value);
}
