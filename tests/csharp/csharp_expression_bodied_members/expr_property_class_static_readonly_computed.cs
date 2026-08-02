// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_static_readonly_computed
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class Consts { public static int Ten => 10; }
__Check((Consts.Ten).ToString(), "10");
