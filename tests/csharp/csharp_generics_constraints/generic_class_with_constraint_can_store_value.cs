// vybe-test: csharp/csharp_generics_constraints/generic_class_with_constraint_can_store_value
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

using static __Harness;

var holder = new Holder<string> { Value = "abc" }
;
__P((holder.Value).ToString());
__Check("abc");

class Holder<T> where T : class { public T Value { get; set; } }

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
