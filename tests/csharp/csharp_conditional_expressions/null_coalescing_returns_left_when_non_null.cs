// vybe-test: csharp/csharp_conditional_expressions/null_coalescing_returns_left_when_non_null
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s="hello";
__Check((s??"default").ToString(), "hello");
