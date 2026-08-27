// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

using static __Harness;

// dictionary_enumeration_order
var tuple = (left: 35, right: 36);
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
