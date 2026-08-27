// vybe-test: csharp/csharp_linq_advanced/reverse_inverts_order_of_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

using static __Harness;

var result=new[]{1,2,3}
.Reverse();
foreach(var n in result) __P((n).ToString());
__Check("3\n2\n1");

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
