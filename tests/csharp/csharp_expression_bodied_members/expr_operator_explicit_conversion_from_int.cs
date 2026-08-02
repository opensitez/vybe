// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_explicit_conversion_from_int
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Wrap { public int V; public static explicit operator Wrap(int n) => new Wrap { V = n }; }
Wrap w = (Wrap)9; __Check((w.V).ToString(), "9");
