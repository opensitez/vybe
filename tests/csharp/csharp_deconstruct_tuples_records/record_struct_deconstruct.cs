// vybe-test: csharp/csharp_deconstruct_tuples_records/record_struct_deconstruct
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Vec(int X,int Y); var (x,y)=new Vec(8,1); __Check((x-y).ToString(), "7");
