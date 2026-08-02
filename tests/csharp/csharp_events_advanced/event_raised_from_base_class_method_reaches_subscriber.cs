// vybe-test: csharp/csharp_events_advanced/event_raised_from_base_class_method_reaches_subscriber
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class BaseNotifier { public event Action Changed; protected void Raise() { Changed(); } public void Touch() { Raise(); } } class ChildNotifier : BaseNotifier { } var notifier = new ChildNotifier(); notifier.Changed += () => __Check(("changed").ToString(), "changed"); notifier.Touch();
