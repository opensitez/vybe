// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_struct_deconstruct
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Pair{public int X,Y; public void Deconstruct(out int x,out int y){x=X;y=Y;}} var (x,y)=new Pair{X=4,Y=6}; __Check((x*y).ToString(), "24");
