// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_deconstruct_four
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (a,b,c,d)=(1,2,3,4); __Check((a+b+c+d).ToString(), "10");
