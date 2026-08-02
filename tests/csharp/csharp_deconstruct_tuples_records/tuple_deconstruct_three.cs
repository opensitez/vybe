// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_deconstruct_three
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (a,b,c)=(1,2,3); __Check((a+b+c).ToString(), "6");
