// vybe-test: csharp/csharp_events_advanced/event_lambda_captures_outer_local_state
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

int count = 0;
var alarm = new Alarm();
alarm.Triggered += () => { count++; __P((count).ToString()); }
;
alarm.Fire();
alarm.Fire();
__Check("1\n2");

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
