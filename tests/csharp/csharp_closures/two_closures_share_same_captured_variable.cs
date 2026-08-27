// vybe-test: csharp/csharp_closures/two_closures_share_same_captured_variable
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

using static __Harness;

int shared = 0;
System.Action add = () => shared++;
System.Func<int> read = () => shared;
add();
add();
__P((read()).ToString());
__Check("2");

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
