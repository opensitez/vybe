// vybe-test: csharp/csharp_static_classes/static_constructor_runs_once_before_first_member_access
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

using static __Harness;

__P((Singleton.Value).ToString());
__P((Singleton.InitCount).ToString());
__Check("42\n1");

class Singleton {
    public static int InitCount = 0;
    static Singleton() { InitCount++; }
    public static int Value = 42;
}

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
