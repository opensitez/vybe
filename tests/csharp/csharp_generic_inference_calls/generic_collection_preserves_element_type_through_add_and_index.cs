// vybe-test: csharp/csharp_generic_inference_calls/generic_collection_preserves_element_type_through_add_and_index
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var scores = new Dictionary<string, int>();
scores.Add("ada", 99);
scores.Add("lin", 88);
__Check((scores["ada"]).ToString(), "99");
__Check((scores.ContainsKey("lin")).ToString(), "True");
