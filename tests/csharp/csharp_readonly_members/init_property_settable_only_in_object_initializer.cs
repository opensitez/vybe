// vybe-test: csharp/csharp_readonly_members/init_property_settable_only_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

using static __Harness;

var c=new Config{Port=443}
;
__P((c.Port).ToString());
__Check("443");

class Config{public int Port{get;init;}=80;}

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
