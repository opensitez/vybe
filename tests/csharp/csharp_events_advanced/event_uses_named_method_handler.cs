// vybe-test: csharp/csharp_events_advanced/event_uses_named_method_handler
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

void OnTriggered() { __P(("ring").ToString()); }
var alarm = new Alarm();
alarm.Triggered += OnTriggered;
alarm.Fire();
__Check("ring");

class Alarm { public event Action Triggered; public void Fire() { Triggered(); } }

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
