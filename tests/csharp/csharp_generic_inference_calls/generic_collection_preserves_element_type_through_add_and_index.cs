// vybe-test: csharp/csharp_generic_inference_calls/generic_collection_preserves_element_type_through_add_and_index
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var scores = new Dictionary<string, int>();
scores.Add("ada", 99);
scores.Add("lin", 88);
__P((scores["ada"]).ToString());
__P((scores.ContainsKey("lin")).ToString());
__Check("99\nTrue");
