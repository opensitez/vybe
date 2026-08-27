// vybe-test: csharp/type_features/lambda_expression_foreach
// origin: languages/csharp/tests/csharp/test_type_features.rs

using static __Harness;

var arr = new int[] { 1, 2, 3, 4, 5 }
;
var sum = 0;
foreach (var x in arr) { sum = sum + x; }
__P((sum).ToString());
__Check("15");

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
