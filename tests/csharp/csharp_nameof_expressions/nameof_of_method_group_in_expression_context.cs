// vybe-test: csharp/csharp_nameof_expressions/nameof_of_method_group_in_expression_context
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Ops{public void Execute(){}} __Check((nameof(Ops)+"."+nameof(Ops.Execute)).ToString(), "Ops.Execute");
