// vybe-test: csharp/csharp_nameof_expressions/nameof_local_int_variable_returns_identifier
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int itemCount=0; __Check((nameof(itemCount)).ToString(), "itemCount");
