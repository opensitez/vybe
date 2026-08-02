// vybe-test: csharp/csharp_nameof_expressions/nameof_local_bool_variable_returns_identifier
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool isReady=true; __Check((nameof(isReady)).ToString(), "isReady");
