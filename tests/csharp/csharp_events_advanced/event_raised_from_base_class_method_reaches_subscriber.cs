// vybe-test: csharp/csharp_events_advanced/event_raised_from_base_class_method_reaches_subscriber
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

using System; class BaseNotifier { public event Action Changed; protected void Raise() { Changed(); } public void Touch() { Raise(); } } class ChildNotifier : BaseNotifier { } var notifier = new ChildNotifier(); notifier.Changed += () => __P(("changed").ToString()); notifier.Touch();
__Check("changed");
