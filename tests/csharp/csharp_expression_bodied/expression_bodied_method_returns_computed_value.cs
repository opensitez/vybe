// vybe-test: csharp/csharp_expression_bodied/expression_bodied_method_returns_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Calc{public int Double(int n)=>n*2;}
__Check((new Calc().Double(5)).ToString(), "10");
