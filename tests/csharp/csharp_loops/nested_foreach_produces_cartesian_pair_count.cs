// vybe-test: csharp/csharp_loops/nested_foreach_produces_cartesian_pair_count
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

using static __Harness;

int count=0;
foreach(var a in new[]{1,2})
    foreach(var b in new[]{1,2,3})
        count++;
__P((count).ToString());
__Check("6");

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
