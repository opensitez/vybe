// vybe-test: csharp/csharp_reflection_emit/type_is_assignable_from_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

using static __Harness;

__P((typeof(A).IsAssignableFrom(typeof(B))).ToString());
__P((typeof(B).IsAssignableFrom(typeof(A))).ToString());
__Check("True\nFalse");

class A{}

class B:A{}

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
