// vybe-test: csharp/csharp_collections/list_sort_and_reverse
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<int>();
list.Add(3);
list.Add(1);
list.Add(4);
list.Add(1);
list.Add(5);
list.Sort();
__P((list[0]).ToString());
__P((list[4]).ToString());
list.Reverse();
__P((list[0]).ToString());
__Check("1\n5\n5");

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
