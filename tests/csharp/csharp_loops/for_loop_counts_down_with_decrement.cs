// vybe-test: csharp/csharp_loops/for_loop_counts_down_with_decrement
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

using static __Harness;

string r="";
for(int i=3;i>=1;i--) r+=i;
__P((r).ToString());
__Check("321");

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
