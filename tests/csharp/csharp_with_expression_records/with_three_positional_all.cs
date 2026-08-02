// vybe-test: csharp/csharp_with_expression_records/with_three_positional_all
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Triple(int A,int B,int C); var u=(new Triple(1,2,3)) with{A=4,B=5,C=6}; __Check((u.A+u.B+u.C).ToString(), "15");
