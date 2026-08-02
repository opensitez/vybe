// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_from_method
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ValueTuple<int,int> Pair()=>(4,5); var (x,y)=Pair(); __Check((x+y).ToString(), "9");
