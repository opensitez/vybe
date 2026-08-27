// vybe-test: csharp/csharp_generics_constraints/generic_method_can_use_list_of_t_argument
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;
using System.Collections.Generic;

int Count<T>(List<T> items) { return items.Count; }
__P((Count(new List<string> { "a", "b" })).ToString());
__Check("2");

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
