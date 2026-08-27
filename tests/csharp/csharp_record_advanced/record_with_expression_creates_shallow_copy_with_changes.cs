// vybe-test: csharp/csharp_record_advanced/record_with_expression_creates_shallow_copy_with_changes
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

using static __Harness;

var c1=new Config(80,"localhost");
var c2=c1 with{Port=443}
;
__P((c1.Port).ToString());
__P((c2.Port).ToString());
__P((c2.Host).ToString());
__Check("80\n443\nlocalhost");

record Config(int Port,string Host);

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
