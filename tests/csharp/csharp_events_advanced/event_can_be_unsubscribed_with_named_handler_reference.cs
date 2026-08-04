// vybe-test: csharp/csharp_events_advanced/event_can_be_unsubscribed_with_named_handler_reference
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

using System; class Bus { public event Action Arrived; public void Fire() { if (Arrived != null) Arrived(); } } void OnArrived() { __P(("arrived").ToString()); } var bus = new Bus(); bus.Arrived += OnArrived; bus.Arrived -= OnArrived; bus.Fire(); __P(("done").ToString());
__Check("done");
