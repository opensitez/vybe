// vybe-test: csharp/csharp_expression_bodied_members/expr_constructor_struct_expression_body_single_field
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Id { public int Value; public Id(int v) => Value = v; }
__Check((new Id(42).Value).ToString(), "42");
