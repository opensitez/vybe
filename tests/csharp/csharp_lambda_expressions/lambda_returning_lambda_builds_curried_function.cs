// vybe-test: csharp/csharp_lambda_expressions/lambda_returning_lambda_builds_curried_function
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,System.Func<int,int>> add = a => b => a+b;
var add5 = add(5);
__Check((add5(3)).ToString(), "8");
