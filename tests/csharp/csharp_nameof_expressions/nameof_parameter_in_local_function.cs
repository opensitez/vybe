// vybe-test: csharp/csharp_nameof_expressions/nameof_parameter_in_local_function
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Outer(){void Inner(int offset){__Check((nameof(offset)).ToString(), "offset");} Inner(3);} Outer();
