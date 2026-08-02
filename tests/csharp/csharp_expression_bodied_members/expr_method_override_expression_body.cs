// vybe-test: csharp/csharp_expression_bodied_members/expr_method_override_expression_body
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public virtual int Id() => 1; }
class Derived : Base { public override int Id() => 2; }
__Check((new Derived().Id()).ToString(), "2");
