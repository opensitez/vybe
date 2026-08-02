// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_implicit_conversion_to_int
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Wrap { public int V; public static implicit operator int(Wrap w) => w.V; }
Wrap w = new Wrap { V = 42 }; int n = w; __Check((n).ToString(), "42");
