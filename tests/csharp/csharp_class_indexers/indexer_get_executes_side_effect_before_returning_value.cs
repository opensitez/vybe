// vybe-test: csharp/csharp_class_indexers/indexer_get_executes_side_effect_before_returning_value
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

var item = new CustomLogger();
item[0] = "Hit";
__P(item[0]);
__Check("Hit");

class CustomLogger {
    public string hits = "";
    public string this[int i] {
        get => hits;
        set => hits = value;
    }
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
