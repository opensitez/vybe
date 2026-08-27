// vybe-test: csharp/csharp_design_patterns/singleton_returns_same_instance_on_repeated_calls
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

using static __Harness;

Singleton.Instance.Val=42;
__P((Singleton.Instance.Val).ToString());
__Check("42");

class Singleton{
    static Singleton _inst;
    public int Val;
    public static Singleton Instance=>_inst??=new Singleton();
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
