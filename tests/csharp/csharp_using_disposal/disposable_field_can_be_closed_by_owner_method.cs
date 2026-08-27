// vybe-test: csharp/csharp_using_disposal/disposable_field_can_be_closed_by_owner_method
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;
using System;

new Owner().Close();
__Check("disposed");

class Resource : IDisposable { public void Dispose() { __P(("disposed").ToString()); } }

class Owner { Resource resource = new Resource(); public void Close() { resource.Dispose(); } }

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
