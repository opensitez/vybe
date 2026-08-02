// vybe-test: csharp/csharp_expression_bodied_members/expr_method_nested_class_delegates_to_outer
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer { public int Base => 10; public class Inner { Outer o; public Inner(Outer owner) { o = owner; } public int Boost() => o.Base + 5; } }
__Check((new Outer.Inner(new Outer()).Boost()).ToString(), "15");
