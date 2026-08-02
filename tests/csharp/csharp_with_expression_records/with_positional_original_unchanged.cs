// vybe-test: csharp/csharp_with_expression_records/with_positional_original_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X,int Y); var p=new Point(1,2); var q=p with{X=9}; __Check((p.X).ToString(), "1");
