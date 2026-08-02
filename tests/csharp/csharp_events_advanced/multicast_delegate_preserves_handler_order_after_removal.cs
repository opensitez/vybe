// vybe-test: csharp/csharp_events_advanced/multicast_delegate_preserves_handler_order_after_removal
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using System; void A() { Console.WriteLine("A"); } void B() { Console.WriteLine("B"); } void C() { Console.WriteLine("C"); } Action action = A; action += B; action += C; action -= B; action();
