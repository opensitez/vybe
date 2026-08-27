// vybe-test: csharp/csharp_delegates/func_delegate
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

using static __Harness;

Func<int, int> square = x => x * x;
__P((square(5)).ToString());
__P((square(8)).ToString());
__Check("25\n64");

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
