// vybe-test: csharp/csharp_nameof_expressions/nameof_var_inferred_local_returns_identifier
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var delta=1; __Check((nameof(delta)).ToString(), "delta");
