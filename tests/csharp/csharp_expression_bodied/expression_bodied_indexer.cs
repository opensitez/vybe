// vybe-test: csharp/csharp_expression_bodied/expression_bodied_indexer
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

using static __Harness;

__P((new Bag()[2]).ToString());
__Check("3");

class Bag{int[]data={1,2,3};public int this[int i]=>data[i];}

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
