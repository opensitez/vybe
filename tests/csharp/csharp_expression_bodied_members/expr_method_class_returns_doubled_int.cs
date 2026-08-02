// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_returns_doubled_int
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Calc { public int Double(int n) => n * 2; }
__Check((new Calc().Double(5)).ToString(), "10");
