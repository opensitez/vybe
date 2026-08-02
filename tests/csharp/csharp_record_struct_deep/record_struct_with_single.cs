// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_single
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Point(int X,int Y); var p=new Point(1,2); var q=p with{X=9}; __Check((p.X).ToString(), "1"); __Check((q.X).ToString(), "9");
