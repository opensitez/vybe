// vybe-test: csharp/csharp_events_advanced/multicast_delegate_preserves_handler_order_after_removal
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

void A() { __P(("A").ToString()); }
void B() { __P(("B").ToString()); }
void C() { __P(("C").ToString()); }
Action action = A;
action += B;
action += C;
action -= B;
action();
__Check("A\nC");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
