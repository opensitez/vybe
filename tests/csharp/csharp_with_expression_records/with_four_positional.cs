// vybe-test: csharp/csharp_with_expression_records/with_four_positional
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Quad(int A,int B,int C,int D); var r=(new Quad(1,2,3,4)) with{D=10}; __Check((r.D).ToString(), "10");
