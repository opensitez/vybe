// vybe-test: csharp/csharp_events_advanced/event_lambda_captures_outer_local_state
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

using System; class Alarm { public event Action Triggered; public void Fire() { Triggered(); } } int count = 0; var alarm = new Alarm(); alarm.Triggered += () => { count++; __P((count).ToString()); }; alarm.Fire(); alarm.Fire();
__Check("1\n2");
