// vybe-test: csharp/csharp_events_advanced/multicast_delegate_combines_two_named_handlers
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using static __Harness;
using System;

void First() { __P(("A").ToString()); }
void Second() { __P(("B").ToString()); }
Action action = First;
action += Second;
action();
__Check("A\nB");

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
