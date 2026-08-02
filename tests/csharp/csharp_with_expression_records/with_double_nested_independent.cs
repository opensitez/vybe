// vybe-test: csharp/csharp_with_expression_records/with_double_nested_independent
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Pair(int A,int B); var p=new Pair(1,1); var a=p with{A=2}; var b=p with{B=3}; __Check((a.A).ToString(), "2"); __Check((b.B).ToString(), "3");
