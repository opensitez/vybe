// vybe-test: csharp/csharp_events_advanced/event_subscription_inside_constructor_is_triggered_by_method
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

var sensor = new Sensor();
sensor.Fire();
__Check("armed");

class Sensor { public event Action Triggered; public Sensor() { Triggered += () => __P(("armed").ToString()); } public void Fire() { Triggered(); } }

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
