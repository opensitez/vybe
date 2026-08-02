// vybe-test: csharp/csharp_conditional_expressions/conditional_expression_in_argument_position
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=7;
__Check((string.Format("{0}",n%2==0?"even":"odd")).ToString(), "odd");
