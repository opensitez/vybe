// vybe-test: csharp/csharp_events_advanced/event_can_be_unsubscribed_with_named_handler_reference
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

void OnArrived() { __P(("arrived").ToString()); }
var bus = new Bus();
bus.Arrived += OnArrived;
bus.Arrived -= OnArrived;
bus.Fire();
__P(("done").ToString());
__Check("done");

class Bus { public event Action Arrived; public void Fire() { if (Arrived != null) Arrived(); } }

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
