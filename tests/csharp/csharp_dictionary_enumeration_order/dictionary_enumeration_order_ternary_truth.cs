// vybe-test: csharp/csharp_dictionary_enumeration_order/dictionary_enumeration_order_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_enumeration_order.rs

using static __Harness;

// dictionary_enumeration_order
int seed = 35;
bool cond = seed % 2 == 0;
__P((cond || !cond).ToString());
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
