// vybe-test: csharp/csharp_lambda_expressions/lambda_passed_as_argument_to_higher_order_method
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Apply(System.Func<int,int,int> op, int a, int b) => op(a,b);
__Check((Apply((a,b) => a+b, 3, 4)).ToString(), "7");
