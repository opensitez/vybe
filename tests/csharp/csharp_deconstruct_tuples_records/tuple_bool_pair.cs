// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_bool_pair
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (on,flag)=(true,false); __Check((on).ToString(), "True"); __Check((flag).ToString(), "False");
