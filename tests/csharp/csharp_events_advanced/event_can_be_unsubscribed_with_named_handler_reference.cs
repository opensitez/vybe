// vybe-test: csharp/csharp_events_advanced/event_can_be_unsubscribed_with_named_handler_reference
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using System; class Bus { public event Action Arrived; public void Fire() { if (Arrived != null) Arrived(); } } void OnArrived() { Console.WriteLine("arrived"); } var bus = new Bus(); bus.Arrived += OnArrived; bus.Arrived -= OnArrived; bus.Fire(); Console.WriteLine("done");
