// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

using static __Harness;

// nullable_value_operators
var set = new System.Collections.Generic.HashSet<int>();
set.Add(57);
set.Add(57);
__P((set.Count == 1).ToString());
__Check("True");

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
