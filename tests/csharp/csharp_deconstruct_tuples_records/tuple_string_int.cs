// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_string_int
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (name,n)=("Ada",42); __Check((name).ToString(), "Ada"); __Check((n).ToString(), "42");
