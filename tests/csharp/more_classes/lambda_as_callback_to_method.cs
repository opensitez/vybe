// vybe-test: csharp/more_classes/lambda_as_callback_to_method
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

var u = new Util();
__P((u.Apply(21)).ToString());
__Check("42");

class Util {
            public int Apply(int x) {
                return x * 2;
            }
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
