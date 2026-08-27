// vybe-test: csharp/oop_advanced/static_class_with_constants
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

__P((Constants.Pi).ToString());
__P((Constants.MaxSize).ToString());
__Check("3.14159\n100");

static class Constants {
    public const double Pi = 3.14159;
    public const int MaxSize = 100;
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
