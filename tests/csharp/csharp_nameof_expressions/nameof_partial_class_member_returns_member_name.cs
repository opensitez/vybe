// vybe-test: csharp/csharp_nameof_expressions/nameof_partial_class_member_returns_member_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Partial{public int Id;} __Check((nameof(Partial.Id)).ToString(), "Id");
