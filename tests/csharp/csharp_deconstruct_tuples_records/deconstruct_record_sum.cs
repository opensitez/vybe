// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_record_sum
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record V(int A,int B,int C); var (a,b,c)=new V(1,2,3); __Check((a+b+c).ToString(), "6");
