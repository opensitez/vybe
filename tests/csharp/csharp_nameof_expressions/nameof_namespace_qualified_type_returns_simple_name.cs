// vybe-test: csharp/csharp_nameof_expressions/nameof_namespace_qualified_type_returns_simple_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((nameof(System.String)).ToString(), "String");
