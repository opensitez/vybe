// vybe-test: csharp/csharp_using_disposal/disposable_can_accumulate_state_before_disposal
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using static __Harness;
using System;

using (var buffer = new Buffer()) { buffer.Add(); buffer.Add(); }
__Check("2");

class Buffer : IDisposable { int count; public void Add() { count++; } public void Dispose() { __P((count).ToString()); } }

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
