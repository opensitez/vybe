// vybe-test: csharp/csharp_nameof_expressions/nameof_typeof_int_expression_returns_keyword_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((nameof(int)).ToString(), "int");
