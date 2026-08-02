// vybe-test: csharp/csharp_with_expression_records/with_preserves_other_nominal
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Pair{public int A{get;init;} public int B{get;init;}} var q=(new Pair{A=1,B=2}) with{A=9}; __Check((q.B).ToString(), "2");
