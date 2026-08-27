// vybe-test: csharp/csharp_linq_expressions_lambda_compilation/expression_compile_case_19

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

System.Linq.Expressions.Expression<Func<int, int>> expr = x => x * 2;
var func = expr.Compile();
int res = func(19);
__P(res.ToString());
__Check("38");
