// vybe-test: csharp/csharp_collections/list_indexof
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<string>();
list.Add("a");
list.Add("b");
list.Add("c");
__P((list.IndexOf("b")).ToString());
__P((list.IndexOf("z")).ToString());
__Check("1\n-1");

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
