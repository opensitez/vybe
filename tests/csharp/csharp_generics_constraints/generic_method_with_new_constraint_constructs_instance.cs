// vybe-test: csharp/csharp_generics_constraints/generic_method_with_new_constraint_constructs_instance
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

T Create<T>() where T : new() { return new T(); }
__P((Create<Box>().Value).ToString());
__Check("9");

class Box { public int Value = 9; }

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
