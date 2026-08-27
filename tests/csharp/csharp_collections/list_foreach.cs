// vybe-test: csharp/csharp_collections/list_foreach
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<int>();
list.Add(1);
list.Add(2);
list.Add(3);
int sum = 0;
foreach (var x in list) {
    sum += x;
}
__P((sum).ToString());
__Check("6");

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
