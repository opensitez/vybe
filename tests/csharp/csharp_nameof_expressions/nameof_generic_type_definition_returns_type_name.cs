// vybe-test: csharp/csharp_nameof_expressions/nameof_generic_type_definition_returns_type_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T>{public T Item;} __Check((nameof(Box)).ToString(), "Box");
