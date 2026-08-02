// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_uses_field_in_body
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Scale { public int factor = 3; public int Apply(int n) => n * factor; }
__Check((new Scale().Apply(4)).ToString(), "12");
