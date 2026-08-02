// vybe-test: csharp/csharp_with_expression_records/with_chained_three_steps
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Box(int W,int H,int D); var a=new Box(1,2,3); var b=a with{W=4}; var c=b with{H=5}; var d=c with{D=6}; __Check((a.W).ToString(), "1"); __Check((d.W).ToString(), "4"); __Check((d.H).ToString(), "5"); __Check((d.D).ToString(), "6");
