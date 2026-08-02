// vybe-test: csharp/csharp_deconstruct_tuples_records/readonly_record_struct_deconstruct
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly record struct Pair(int A,int B); var (a,b)=new Pair(2,5); __Check((a*b).ToString(), "10");
