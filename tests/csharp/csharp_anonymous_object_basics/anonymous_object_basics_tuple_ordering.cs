// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

using static __Harness;

// anonymous_object_basics
var tuple = (left: 38, right: 39);
__P((tuple.left < tuple.right).ToString());
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
