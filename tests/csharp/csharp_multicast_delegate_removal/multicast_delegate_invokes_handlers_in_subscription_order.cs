// vybe-test: csharp/csharp_multicast_delegate_removal/multicast_delegate_invokes_handlers_in_subscription_order
// origin: languages/csharp/tests/csharp/test_csharp_multicast_delegate_removal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System;
void A() { __Check(("A").ToString(), "A"); }
void B() { __Check(("B").ToString(), "B"); }
Action chain = A;
chain += B;
chain();
