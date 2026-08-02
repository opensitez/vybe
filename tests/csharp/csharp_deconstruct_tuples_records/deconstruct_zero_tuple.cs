// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_zero_tuple
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (a,b)=(0,0); __Check((a==b).ToString(), "True");
