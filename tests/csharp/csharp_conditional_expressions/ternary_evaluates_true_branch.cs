// vybe-test: csharp/csharp_conditional_expressions/ternary_evaluates_true_branch
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=10;
__Check((x>5?"big":"small").ToString(), "big");
