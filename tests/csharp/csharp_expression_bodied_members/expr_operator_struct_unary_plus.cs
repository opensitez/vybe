// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_struct_unary_plus
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Signed { public int V; public static Signed operator +(Signed s) => new Signed { V = +s.V }; }
__Check(((+new Signed { V = 7 }).V).ToString(), "7");
