// vybe-test: csharp/csharp_with_expression/with_expression_on_record_with_init_property
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

using static __Harness;

var base_ = new Config { Host = "localhost", Port = 80 }
;
var prod = base_ with { Port = 443 }
;
__P((prod.Host).ToString());
__P((prod.Port).ToString());
__Check("localhost\n443");

record Config { public string Host { get; init; } public int Port { get; init; } }

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
