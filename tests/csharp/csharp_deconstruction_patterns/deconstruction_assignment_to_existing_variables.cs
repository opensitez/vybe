// vybe-test: csharp/csharp_deconstruction_patterns/deconstruction_assignment_to_existing_variables
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

using static __Harness;

int x=0, y=0;
(x, y) = (5, 10);
__P((x).ToString());
__P((y).ToString());
__Check("5\n10");

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
