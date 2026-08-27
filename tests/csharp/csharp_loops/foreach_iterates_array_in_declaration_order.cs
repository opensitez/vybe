// vybe-test: csharp/csharp_loops/foreach_iterates_array_in_declaration_order
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

using static __Harness;

int s=0;
foreach(var x in new[]{3,1,4}) s+=x;
__P((s).ToString());
__Check("8");

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
