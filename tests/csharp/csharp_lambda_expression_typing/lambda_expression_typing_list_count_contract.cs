// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

using static __Harness;

// lambda_expression_typing
var values = new System.Collections.Generic.List<int> { 76, 77, 76 }
;
__P((values.Count == 3).ToString());
__Check("True");

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
