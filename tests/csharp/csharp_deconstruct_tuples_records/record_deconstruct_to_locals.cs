// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_to_locals
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record R(int A,int B); var r=new R(2,3); int x,y; (x,y)=r; __Check((x*y).ToString(), "6");
