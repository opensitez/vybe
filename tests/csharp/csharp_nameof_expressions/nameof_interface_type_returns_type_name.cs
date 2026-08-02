// vybe-test: csharp/csharp_nameof_expressions/nameof_interface_type_returns_type_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRepository{} __Check((nameof(IRepository)).ToString(), "IRepository");
