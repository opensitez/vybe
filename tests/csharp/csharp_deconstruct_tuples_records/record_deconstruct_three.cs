// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_three
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Triple(int A,int B,int C); var (a,b,c)=new Triple(1,2,3); __Check((a+b+c).ToString(), "6");
