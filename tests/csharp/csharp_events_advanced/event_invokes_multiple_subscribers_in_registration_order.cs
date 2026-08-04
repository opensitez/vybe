// vybe-test: csharp/csharp_events_advanced/event_invokes_multiple_subscribers_in_registration_order
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

using System; class Counter { public event Action Tick; public void Fire() { Tick(); } } var counter = new Counter(); counter.Tick += () => __P(("first").ToString()); counter.Tick += () => __P(("second").ToString()); counter.Fire();
__Check("first\nsecond");
