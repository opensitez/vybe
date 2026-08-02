// vybe-test: csharp/csharp_events_advanced/multicast_delegate_combines_two_named_handlers
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; void First() { __Check(("A").ToString(), "A"); } void Second() { __Check(("B").ToString(), "B"); } Action action = First; action += Second; action();
