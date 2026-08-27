// vybe-test: csharp/csharp_class_indexers/indexer_get_can_compute_derived_value_from_state
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

__P((new Scale()[5]).ToString());
__Check("10");

class Scale {
    int factor = 2;
    public int this[int input] { get { return input * factor; } }
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
