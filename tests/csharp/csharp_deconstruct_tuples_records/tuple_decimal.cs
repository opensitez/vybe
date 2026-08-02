// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_decimal
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (a,b)=(1.5m,2.5m); __Check((a+b).ToString(), "4.0");
