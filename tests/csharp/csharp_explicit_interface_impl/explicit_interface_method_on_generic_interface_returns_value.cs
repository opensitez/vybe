// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_on_generic_interface_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

using static __Harness;

IBox<int> box = new NumberBox();
__P((box.Unwrap()).ToString());
__Check("42");

interface IBox<T> { T Unwrap(); }

class NumberBox : IBox<int> {
    int IBox<int>.Unwrap() { return 42; }
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
