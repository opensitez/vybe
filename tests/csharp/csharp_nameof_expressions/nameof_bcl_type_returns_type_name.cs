// vybe-test: csharp/csharp_nameof_expressions/nameof_bcl_type_returns_type_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((nameof(System.DateTime)).ToString(), "DateTime");
