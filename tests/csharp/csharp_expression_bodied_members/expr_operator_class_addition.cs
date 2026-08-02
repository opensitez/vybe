// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_class_addition
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Num { public int V; public static Num operator +(Num a, Num b) => new Num { V = a.V + b.V }; }
__Check(((new Num { V = 3 } + new Num { V = 4 }).V).ToString(), "7");
