// vybe-test: csharp/csharp_with_expression_records/with_two_branches_same_source
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Pair(int A,int B); var p=new Pair(1,2); var x=p with{A=9}; var y=p with{B=8}; __Check((x.A).ToString(), "9"); __Check((y.B).ToString(), "8");
