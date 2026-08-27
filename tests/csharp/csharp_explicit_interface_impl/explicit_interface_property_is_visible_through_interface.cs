// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_property_is_visible_through_interface
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

using static __Harness;

IValueHolder holder = new Counter();
__P((holder.Value).ToString());
__Check("12");

interface IValueHolder { int Value { get; } }

class Counter : IValueHolder {
    int IValueHolder.Value { get { return 12; } }
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
