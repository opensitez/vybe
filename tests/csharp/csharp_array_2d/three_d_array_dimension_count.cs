// vybe-test: csharp/csharp_array_2d/three_d_array_dimension_count
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

using static __Harness;

int[,,] t=new int[2,3,4];
__P((t.Rank).ToString());
__Check("3");

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
