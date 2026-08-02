// vybe-test: csharp/csharp_events_advanced/event_remove_handler_stops_future_invocation
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using System; class Counter { public event Action Tick; public void Fire() { Tick(); } } var counter = new Counter(); Action handler = () => Console.WriteLine("kept"); counter.Tick += handler; counter.Tick += () => Console.WriteLine("temp"); counter.Tick -= handler; counter.Fire();
