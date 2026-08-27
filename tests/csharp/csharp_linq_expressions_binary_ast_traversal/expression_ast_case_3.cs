// vybe-test: csharp/csharp_linq_expressions_binary_ast_traversal/expression_ast_case_3

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

System.Linq.Expressions.Expression<Func<int, bool>> expr = x => x > 10;
var binary = expr.Body as System.Linq.Expressions.BinaryExpression;
__P(binary.NodeType.ToString());
__Check("GreaterThan");
