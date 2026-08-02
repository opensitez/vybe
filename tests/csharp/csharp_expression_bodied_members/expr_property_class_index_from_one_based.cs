// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_index_from_one_based
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Row { public int Index = 0; public int Display => Index + 1; }
__Check((new Row { Index = 4 }.Display).ToString(), "5");
