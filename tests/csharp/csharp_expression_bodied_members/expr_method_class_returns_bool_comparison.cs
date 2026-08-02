// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_returns_bool_comparison
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Check { public bool IsZero(int n) => n == 0; }
__Check((new Check().IsZero(0)).ToString(), "True"); __Check((new Check().IsZero(1)).ToString(), "False");
