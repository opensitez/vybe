// vybe-test: csharp/csharp_readonly_members/readonly_static_field_initialized_at_class_load
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

using static __Harness;

__P((Config.Env).ToString());
__Check("prod");

class Config{public static readonly string Env="prod";}

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
