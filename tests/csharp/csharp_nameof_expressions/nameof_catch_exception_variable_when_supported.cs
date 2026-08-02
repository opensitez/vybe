// vybe-test: csharp/csharp_nameof_expressions/nameof_catch_exception_variable_when_supported
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try{throw new System.Exception("x");}catch(System.Exception ex){__Check((nameof(ex)).ToString(), "ex");}
