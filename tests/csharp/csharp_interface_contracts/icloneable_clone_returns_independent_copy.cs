// vybe-test: csharp/csharp_interface_contracts/icloneable_clone_returns_independent_copy
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

using static __Harness;

var original = new Box { Value=5 }
;
var copy = (Box)original.Clone();
copy.Value = 99;
__P((original.Value).ToString());
__Check("5");

class Box : System.ICloneable {
    public int Value;
    public object Clone() => new Box { Value = Value };
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
