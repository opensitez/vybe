// vybe-test: csharp/csharp_collections/list_add_and_count
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<int>();
list.Add(10);
list.Add(20);
list.Add(30);
__P((list.Count).ToString());
__P((list[1]).ToString());
__Check("3\n20");

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
