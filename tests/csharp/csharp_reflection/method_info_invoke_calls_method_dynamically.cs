// vybe-test: csharp/csharp_reflection/method_info_invoke_calls_method_dynamically
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

using static __Harness;

var obj = new Calc();
var method = typeof(Calc).GetMethod("Double");
__P((method.Invoke(obj, new object[]{5})).ToString());
__Check("10");

class Calc { public int Double(int n) => n * 2; }

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
