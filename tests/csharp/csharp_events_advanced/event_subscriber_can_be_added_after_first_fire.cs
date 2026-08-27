// vybe-test: csharp/csharp_events_advanced/event_subscriber_can_be_added_after_first_fire
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

var clock = new Clock();
clock.Fire();
clock.Tick += () => __P(("later").ToString());
clock.Fire();
__Check("later");

class Clock { public event Action Tick; public void Fire() { if (Tick != null) Tick(); } }

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
