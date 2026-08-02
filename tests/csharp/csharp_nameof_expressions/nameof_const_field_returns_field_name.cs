// vybe-test: csharp/csharp_nameof_expressions/nameof_const_field_returns_field_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Limits{public const int Max=100;} __Check((nameof(Limits.Max)).ToString(), "Max");
