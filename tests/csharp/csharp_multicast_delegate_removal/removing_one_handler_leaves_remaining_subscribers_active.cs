// vybe-test: csharp/csharp_multicast_delegate_removal/removing_one_handler_leaves_remaining_subscribers_active
// origin: languages/csharp/tests/csharp/test_csharp_multicast_delegate_removal.rs

using System;
void A() { Console.WriteLine("A"); }
void B() { Console.WriteLine("B"); }
Action chain = A;
chain += B;
chain -= A;
chain();
