// vybe-test: csharp/csharp_lambda_expressions/lambda_implicitly_typed_with_var_in_local_variable
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var f = (int x) => x + 1;
__Check((f(9)).ToString(), "10");
