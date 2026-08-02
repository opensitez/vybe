// vybe-test: csharp/csharp_with_expression_records/with_positional_single_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X,int Y); var q=(new Point(1,2)) with{X=10}; __Check((q.X).ToString(), "10"); __Check((q.Y).ToString(), "2");
