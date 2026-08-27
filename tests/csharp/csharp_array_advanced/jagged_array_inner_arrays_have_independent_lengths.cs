// vybe-test: csharp/csharp_array_advanced/jagged_array_inner_arrays_have_independent_lengths
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

using static __Harness;

int[][] j=new int[3][];
j[0]=new int[1];
j[1]=new int[2];
j[2]=new int[3];
__P((j[0].Length).ToString());
__P((j[2].Length).ToString());
__Check("1\n3");

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
