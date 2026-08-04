// vybe-test: csharp/csharp_events_advanced/event_subscription_inside_constructor_is_triggered_by_method
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

using System; class Sensor { public event Action Triggered; public Sensor() { Triggered += () => __P(("armed").ToString()); } public void Fire() { Triggered(); } } var sensor = new Sensor(); sensor.Fire();
__Check("armed");
