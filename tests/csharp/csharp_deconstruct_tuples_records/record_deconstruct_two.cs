// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_two
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X,int Y); var (x,y)=new Point(3,4); __Check((x).ToString(), "3"); __Check((y).ToString(), "4");
