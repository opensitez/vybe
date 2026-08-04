// vybe-test: csharp/csharp_events/null_event_invocation_via_conditional_access_is_safe
// origin: languages/csharp/tests/csharp/test_csharp_events.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Source { public event System.Action Fired; public void Fire() => Fired?.Invoke(); }
var s = new Source();
s.Fire();
__P(("ok").ToString());
__Check("ok");
