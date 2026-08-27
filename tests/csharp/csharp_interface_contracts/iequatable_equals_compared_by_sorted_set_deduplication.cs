// vybe-test: csharp/csharp_interface_contracts/iequatable_equals_compared_by_sorted_set_deduplication
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

using static __Harness;

var set = new System.Collections.Generic.HashSet<Id>(
    System.Collections.Generic.EqualityComparer<Id>.Default);
set.Add(new Id{Value=1});
set.Add(new Id{Value=1});
__P((set.Count).ToString());
__Check("1");

class Id : System.IEquatable<Id> {
    public int Value;
    public bool Equals(Id other) => other?.Value == Value;
    public override bool Equals(object o) => o is Id i && Equals(i);
    public override int GetHashCode() => Value;
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
