// vybe-test: csharp/csharp_events_advanced/generic_eventhandler_passes_integer_payload
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

var stream = new Stream();
stream.Produced += (sender, args) => __P((args.Value * 2).ToString());
stream.Emit(6);
__Check("12");

class NumberEventArgs : EventArgs { public int Value { get; set; } }

class Stream { public event EventHandler<NumberEventArgs> Produced; public void Emit(int value) { Produced(this, new NumberEventArgs { Value = value }); } }

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
