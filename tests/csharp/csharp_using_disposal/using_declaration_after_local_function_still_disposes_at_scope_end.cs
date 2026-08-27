// vybe-test: csharp/csharp_using_disposal/using_declaration_after_local_function_still_disposes_at_scope_end
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;
using System;

string Read() { using var resource = new Resource(); return "ok"; }
__P((Read()).ToString());
__Check("disposed\nok");

class Resource : IDisposable { public void Dispose() { __P(("disposed").ToString()); } }

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
