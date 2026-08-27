// vybe-test: csharp/csharp_interface_explicit_impl/class_with_both_explicit_and_public_overloads_picks_by_static_type
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

using static __Harness;

var w = new Widget();
IDescribe i = w;
__P((w.Describe()).ToString());
__P((i.Describe()).ToString());
__Check("widget\ninterface:widget");

interface IDescribe { string Describe(); }

class Widget : IDescribe {
    public string Describe() => "widget";
    string IDescribe.Describe() => "interface:widget";
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
