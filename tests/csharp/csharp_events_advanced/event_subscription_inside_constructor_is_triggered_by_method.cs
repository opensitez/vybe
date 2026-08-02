// vybe-test: csharp/csharp_events_advanced/event_subscription_inside_constructor_is_triggered_by_method
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Sensor { public event Action Triggered; public Sensor() { Triggered += () => __Check(("armed").ToString(), "armed"); } public void Fire() { Triggered(); } } var sensor = new Sensor(); sensor.Fire();
