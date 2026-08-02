// vybe-test: csharp/csharp_events_advanced/event_uses_named_method_handler
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Alarm { public event Action Triggered; public void Fire() { Triggered(); } } void OnTriggered() { __Check(("ring").ToString(), "ring"); } var alarm = new Alarm(); alarm.Triggered += OnTriggered; alarm.Fire();
