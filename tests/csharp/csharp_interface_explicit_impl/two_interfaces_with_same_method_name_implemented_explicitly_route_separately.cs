// vybe-test: csharp/csharp_interface_explicit_impl/two_interfaces_with_same_method_name_implemented_explicitly_route_separately
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

using static __Harness;

ILeft  l = new Both();
IRight r = new Both();
__P((l.Side()).ToString());
__P((r.Side()).ToString());
__Check("left\nright");

interface ILeft  { string Side(); }

interface IRight { string Side(); }

class Both : ILeft, IRight {
    string ILeft.Side()  => "left";
    string IRight.Side() => "right";
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
