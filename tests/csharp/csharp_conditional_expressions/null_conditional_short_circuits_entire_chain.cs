// vybe-test: csharp/csharp_conditional_expressions/null_conditional_short_circuits_entire_chain
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s=null;
__Check((s?.ToUpper()??"nil").ToString(), "nil");
