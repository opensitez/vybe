// vybe-test: csharp/csharp_static_classes/static_constructor_not_re_run_on_second_access
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

using static __Harness;

Registry.Touch();
Registry.Touch();
__P((Registry.Boot).ToString());
__Check("1");

class Registry {
    public static int Boot = 0;
    static Registry() { Boot++; }
    public static void Touch() { }
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
