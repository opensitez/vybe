// vybe-test: csharp/csharp_object_equality/overridden_equals_on_class_reflects_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

using static __Harness;

var x = new Money { Amount = 5 }
;
var y = new Money { Amount = 5 }
;
__P((x.Equals(y)).ToString());
__P((object.ReferenceEquals(x, y)).ToString());
__Check("True\nFalse");

class Money {
    public int Amount;
    public override bool Equals(object obj) =>
        obj is Money m && m.Amount == Amount;
    public override int GetHashCode() => Amount;
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
