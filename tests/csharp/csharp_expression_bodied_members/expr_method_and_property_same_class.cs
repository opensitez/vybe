// vybe-test: csharp/csharp_expression_bodied_members/expr_method_and_property_same_class
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget { public int baseVal = 5; public int Base => baseVal; public int Twice() => Base * 2; }
__Check((new Widget().Twice()).ToString(), "10");
