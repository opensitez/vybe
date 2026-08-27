// vybe-test: csharp/csharp_linq_aggregate_element/min_by_max_by_same_sequence_lengths
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

var words=new[]{"go","stop","run"}
;
__P((words.MinBy(w=>w.Length).Length).ToString());
__P((words.MaxBy(w=>w.Length).Length).ToString());
__Check("2\n4");

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
