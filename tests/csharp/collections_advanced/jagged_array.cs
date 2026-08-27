// vybe-test: csharp/collections_advanced/jagged_array
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

int[][] jagged = new int[3][];
jagged[0] = new int[] { 1, 2 }
;
jagged[1] = new int[] { 3, 4, 5 }
;
jagged[2] = new int[] { 6 }
;
__P((jagged[0].Length).ToString());
__P((jagged[1].Length).ToString());
__P((jagged[1][2]).ToString());
__Check("2\n3\n5");

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
