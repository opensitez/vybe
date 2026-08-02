// vybe-test: csharp/csharp_expression_bodied_members/expr_method_struct_instance_on_stack
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Counter { public int n; public int Next() => ++n; }
var c = new Counter();
__Check((c.Next()).ToString(), "1"); __Check((c.Next()).ToString(), "2");
