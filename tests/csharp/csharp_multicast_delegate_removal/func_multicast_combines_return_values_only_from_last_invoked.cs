// vybe-test: csharp/csharp_multicast_delegate_removal/func_multicast_combines_return_values_only_from_last_invoked
// origin: languages/csharp/tests/csharp/test_csharp_multicast_delegate_removal.rs

using static __Harness;
using System;

Func<int> first = () => { __P(("1").ToString()); return 1; }
;
Func<int> second = () => { __P(("2").ToString()); return 2; }
;
Func<int> chain = first;
chain += second;
__P((chain()).ToString());
__Check("1\n2\n2");

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
