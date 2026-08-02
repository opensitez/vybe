// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.ValueTuple<int,int> Twice(int n)=>(n,n); var (a,b)=Twice(6); __Check((a*b).ToString(), "36");
