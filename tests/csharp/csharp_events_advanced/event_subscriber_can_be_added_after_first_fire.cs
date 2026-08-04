// vybe-test: csharp/csharp_events_advanced/event_subscriber_can_be_added_after_first_fire
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

using System; class Clock { public event Action Tick; public void Fire() { if (Tick != null) Tick(); } } var clock = new Clock(); clock.Fire(); clock.Tick += () => __P(("later").ToString()); clock.Fire();
__Check("later");
