// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_and_method_on_same_struct
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Num { public int V; public static Num operator +(Num a, Num b) => new Num { V = a.V + b.V }; public int Double() => V * 2; }
var n = new Num { V = 3 } + new Num { V = 4 }; __Check((n.Double()).ToString(), "14");
