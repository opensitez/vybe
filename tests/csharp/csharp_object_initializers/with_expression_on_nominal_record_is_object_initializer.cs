// vybe-test: csharp/csharp_object_initializers/with_expression_on_nominal_record_is_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

using static __Harness;

var cfg=new Config() with{Port=443}
;
__P((cfg.Port).ToString());
__Check("443");

record Config{public int Port{get;init;}=80;}

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
