// vybe-test: csharp/csharp_generic_inference_calls/generic_collection_preserves_element_type_through_add_and_index
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;
using System.Collections.Generic;

var scores = new Dictionary<string, int>();
scores.Add("ada", 99);
scores.Add("lin", 88);
__P((scores["ada"]).ToString());
__P((scores.ContainsKey("lin")).ToString());
__Check("99\nTrue");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
