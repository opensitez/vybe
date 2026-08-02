// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_computed_from_state
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Scale { int factor = 2; public int this[int input] => input * factor; }
__Check((new Scale()[5]).ToString(), "10");
