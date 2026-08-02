// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_long_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (lo,hi)=(10000000000L,5L); __Check((lo+hi).ToString(), "10000000005");
