// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_sequential
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (a,b)=(1,2); var (c,d)=(a+b,b); __Check((c).ToString(), "3"); __Check((d).ToString(), "2");
