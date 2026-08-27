// vybe-test: csharp/csharp_deconstruction_patterns/record_positional_deconstruct_extracts_all_fields
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

using static __Harness;

var p = new Point(1,2,3);
var (x,y,z) = p;
__P((x+y+z).ToString());
__Check("6");

record Point(int X, int Y, int Z);

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
