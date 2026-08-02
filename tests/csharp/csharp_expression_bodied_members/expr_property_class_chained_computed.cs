// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_chained_computed
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Chain { public int Base = 2; public int Double => Base * 2; public int Quadruple => Double * 2; }
__Check((new Chain().Quadruple).ToString(), "8");
