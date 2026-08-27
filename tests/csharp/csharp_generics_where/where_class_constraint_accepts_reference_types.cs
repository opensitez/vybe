// vybe-test: csharp/csharp_generics_where/where_class_constraint_accepts_reference_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

using static __Harness;

T Wrap<T>(T v) where T:class=>v;
__P((Wrap("hello")).ToString());
__Check("hello");

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
