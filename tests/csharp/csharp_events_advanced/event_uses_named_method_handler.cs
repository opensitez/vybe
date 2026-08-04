// vybe-test: csharp/csharp_events_advanced/event_uses_named_method_handler
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

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

using System; class Alarm { public event Action Triggered; public void Fire() { Triggered(); } } void OnTriggered() { __P(("ring").ToString()); } var alarm = new Alarm(); alarm.Triggered += OnTriggered; alarm.Fire();
__Check("ring");
