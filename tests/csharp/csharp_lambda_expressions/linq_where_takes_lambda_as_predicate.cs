// vybe-test: csharp/csharp_lambda_expressions/linq_where_takes_lambda_as_predicate
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

using static __Harness;

var evens = new[]{1,2,3,4,5,6}
.Where(n => n%2==0);
__P((evens.Count()).ToString());
__Check("3");

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
