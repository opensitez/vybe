// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_returns_double_ratio
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Ratio { public double Half(double x) => x / 2.0; }
__Check((new Ratio().Half(5.0)).ToString(), "2.5");
