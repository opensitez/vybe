// vybe-test: csharp/interfaces_generics/generic_where_class_constraint
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

var c = new Container<string>();
__P((c.IsNull()).ToString());
c.Value = "hello";
__P((c.IsNull()).ToString());
__Check("True\nFalse");

class Container<T> where T : class {
    public T Value;
    public bool IsNull() { return Value == null; }
}

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
