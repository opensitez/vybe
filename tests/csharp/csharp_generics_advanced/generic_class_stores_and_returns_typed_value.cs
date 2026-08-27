// vybe-test: csharp/csharp_generics_advanced/generic_class_stores_and_returns_typed_value
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

var b = new Box<int> { Value = 42 }
;
__P((b.Value).ToString());
__Check("42");

class Box<T> { public T Value; }

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
