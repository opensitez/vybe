// vybe-test: csharp/basics/array_copy_runtime
// origin: languages/csharp/tests/csharp/test_basics.rs

using static __Harness;

int[] src = new int[] { 10, 20, 30, 40 }
;
int[] dst = new int[] { 0, 0, 0, 0 }
;
Array.Copy(src, dst, 3);
__P((dst[0] + dst[1] + dst[2] + dst[3]).ToString());
__Check("60");

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
