// vybe-test: csharp/csharp_map_set_collections/linked_list_find_returns_matching_node_value
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using static __Harness;
using System.Collections.Generic;

var items = new LinkedList<string>();
items.AddLast("a");
items.AddLast("b");
var node = items.Find("b");
__P((node.Value).ToString());
__Check("b");

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
