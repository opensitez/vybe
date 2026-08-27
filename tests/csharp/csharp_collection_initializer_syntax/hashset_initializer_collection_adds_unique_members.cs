// vybe-test: csharp/csharp_collection_initializer_syntax/hashset_initializer_collection_adds_unique_members
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;
using System.Collections.Generic;

var set = new HashSet<int> { 2, 3, 2 }
;
__P((set.Count).ToString());
__Check("2");

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
