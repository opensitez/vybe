// vybe-test: csharp/csharp_events_advanced/event_remove_handler_stops_future_invocation
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

var counter = new Counter();
Action handler = () => __P(("kept").ToString());
counter.Tick += handler;
counter.Tick += () => __P(("temp").ToString());
counter.Tick -= handler;
counter.Fire();
__Check("temp");

class Counter { public event Action Tick; public void Fire() { Tick(); } }

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
