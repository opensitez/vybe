// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_three_params_product
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Mul3 { public int Prod(int a, int b, int c) => a * b * c; }
__Check((new Mul3().Prod(2, 3, 4)).ToString(), "24");
