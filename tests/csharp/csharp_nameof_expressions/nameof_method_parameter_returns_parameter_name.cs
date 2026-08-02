// vybe-test: csharp/csharp_nameof_expressions/nameof_method_parameter_returns_parameter_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Report(int total){__Check((nameof(total)).ToString(), "total");} Report(1);
