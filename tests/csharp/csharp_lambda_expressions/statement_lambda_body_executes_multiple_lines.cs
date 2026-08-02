// vybe-test: csharp/csharp_lambda_expressions/statement_lambda_body_executes_multiple_lines
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> fact = null;
fact = n => { if(n<=1) return 1; return n*fact(n-1); };
__Check((fact(5)).ToString(), "120");
