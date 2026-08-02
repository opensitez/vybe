// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_nullable_record
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Maybe(int? N); var (n)=new Maybe(5); __Check((n).ToString(), "5");
