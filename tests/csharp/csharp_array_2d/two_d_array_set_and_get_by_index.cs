// vybe-test: csharp/csharp_array_2d/two_d_array_set_and_get_by_index
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

using static __Harness;

int[,] m=new int[3,3];
m[1,2]=99;
__P((m[1,2]).ToString());
__Check("99");

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
