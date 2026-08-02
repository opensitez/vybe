// vybe-test: csharp/csharp_nameof_expressions/nameof_local_function_returns_function_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Compute(){return 1;} __Check((nameof(Compute)).ToString(), "Compute");
