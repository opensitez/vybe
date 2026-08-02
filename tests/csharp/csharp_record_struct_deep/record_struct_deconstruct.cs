// vybe-test: csharp/csharp_record_struct_deep/record_struct_deconstruct
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Vec(int X,int Y); var (x,y)=new Vec(3,4); __Check((x+y).ToString(), "7");
