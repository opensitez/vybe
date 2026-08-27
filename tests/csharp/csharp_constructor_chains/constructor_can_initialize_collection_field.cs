// vybe-test: csharp/csharp_constructor_chains/constructor_can_initialize_collection_field
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;
using System.Collections.Generic;

__P((new Box().Count()).ToString());
__Check("3");

class Box { List<int> values; public Box() { values = new List<int> { 1, 2, 3 }; } public int Count() { return values.Count; } }

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
