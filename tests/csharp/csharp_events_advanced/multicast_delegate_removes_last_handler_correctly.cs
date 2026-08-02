// vybe-test: csharp/csharp_events_advanced/multicast_delegate_removes_last_handler_correctly
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using System; void First() { Console.WriteLine("A"); } void Second() { Console.WriteLine("B"); } Action action = First; action += Second; action -= Second; action();
