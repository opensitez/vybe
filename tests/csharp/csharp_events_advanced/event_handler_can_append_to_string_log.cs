// vybe-test: csharp/csharp_events_advanced/event_handler_can_append_to_string_log
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

string log = "";
var emitter = new Emitter();
emitter.Fired += () => log += "x";
emitter.Fired += () => log += "y";
emitter.Fire();
__P((log).ToString());
__Check("xy");

class Emitter { public event Action Fired; public void Fire() { Fired(); } }

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
