// vybe-test: csharp/csharp_using_disposal/using_statement_disposes_multiple_resources_in_reverse_order
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;
using System;

using (var left = new Resource("left")) using (var right = new Resource("right")) { __P(("body").ToString()); }
__Check("body\nright\nleft");

class Resource : IDisposable { string name; public Resource(string name) { this.name = name; } public void Dispose() { __P((name).ToString()); } }

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
