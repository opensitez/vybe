// vybe-test: csharp/csharp_conditional_expressions/ternary_nested_three_way_comparison
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=0;
__Check((n>0?"pos":n<0?"neg":"zero").ToString(), "zero");
