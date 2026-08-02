// vybe-test: csharp/csharp_events_advanced/event_with_two_subscribers_updates_shared_counter
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Hub { public event Action Ping; public void Fire() { Ping(); } } int total = 0; var hub = new Hub(); hub.Ping += () => total += 2; hub.Ping += () => total += 3; hub.Fire(); __Check((total).ToString(), "5");
