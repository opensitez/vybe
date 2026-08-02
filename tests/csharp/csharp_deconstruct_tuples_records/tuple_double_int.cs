// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_double_int
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (rate,count)=(2.5,4); __Check((rate*count).ToString(), "10");
