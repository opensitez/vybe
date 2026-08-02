// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_two
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Point(int X,int Y); var p=new Point(1,2); var q=p with{X=3,Y=4}; __Check((p.Y).ToString(), "2"); __Check((q.X).ToString(), "3"); __Check((q.Y).ToString(), "4");
