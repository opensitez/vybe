// vybe-test: csharp/csharp_events_advanced/event_invokes_multiple_subscribers_in_registration_order
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Counter { public event Action Tick; public void Fire() { Tick(); } } var counter = new Counter(); counter.Tick += () => __Check(("first").ToString(), "first"); counter.Tick += () => __Check(("second").ToString(), "second"); counter.Fire();
