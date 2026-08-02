// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_discard_first
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (x,_)=(7,8); __Check((x).ToString(), "7");
