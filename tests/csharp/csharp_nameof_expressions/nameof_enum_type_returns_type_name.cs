// vybe-test: csharp/csharp_nameof_expressions/nameof_enum_type_returns_type_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color{Red,Green,Blue} __Check((nameof(Color)).ToString(), "Color");
