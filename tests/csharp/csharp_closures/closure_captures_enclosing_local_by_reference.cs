// vybe-test: csharp/csharp_closures/closure_captures_enclosing_local_by_reference
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

using static __Harness;

int x = 1;
System.Action inc = () => x++;
inc();
inc();
__P((x).ToString());
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
