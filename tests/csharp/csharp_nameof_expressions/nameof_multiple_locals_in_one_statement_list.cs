// vybe-test: csharp/csharp_nameof_expressions/nameof_multiple_locals_in_one_statement_list
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int width=1,height=2; __Check((nameof(width)).ToString(), "width"); __Check((nameof(height)).ToString(), "height");
