// vybe-test: csharp/csharp_generics_constraints/generic_static_field_is_independent_per_closed_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

Counter<int>.Value = 2;
Counter<string>.Value = 5;
__P((Counter<int>.Value).ToString());
__P((Counter<string>.Value).ToString());
__Check("2\n5");

class Counter<T> { public static int Value; }

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
