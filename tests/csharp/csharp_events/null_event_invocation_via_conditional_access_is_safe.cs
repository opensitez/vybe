// vybe-test: csharp/csharp_events/null_event_invocation_via_conditional_access_is_safe
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

using static __Harness;

var s = new Source();
s.Fire();
__P(("ok").ToString());
__Check("ok");

class Source { public event System.Action Fired; public void Fire() => Fired?.Invoke(); }

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
