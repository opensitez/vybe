// vybe-test: csharp/csharp_events_advanced/event_subscriber_can_be_added_after_first_fire
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Clock { public event Action Tick; public void Fire() { if (Tick != null) Tick(); } } var clock = new Clock(); clock.Fire(); clock.Tick += () => __Check(("later").ToString(), "later"); clock.Fire();
