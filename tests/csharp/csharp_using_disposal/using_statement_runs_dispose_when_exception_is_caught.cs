// vybe-test: csharp/csharp_using_disposal/using_statement_runs_dispose_when_exception_is_caught
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;
using System;

try { using (var resource = new Resource()) { __P(("body").ToString()); throw new Exception("boom"); } }
catch (Exception) { __P(("caught").ToString()); }
__Check("body\ndisposed\ncaught");

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
