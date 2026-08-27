// vybe-test: csharp/csharp_lambdas/lambda_with_list_foreach
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

using static __Harness;
using System.Collections.Generic;

var items = new List<int>();
items.Add(1);
items.Add(2);
items.Add(3);
items.ForEach(x => __P((x).ToString()));
__Check("1\n2\n3");

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
