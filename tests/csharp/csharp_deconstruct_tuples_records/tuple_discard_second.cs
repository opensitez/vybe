// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_discard_second
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (_,y)=(99,3); __Check((y).ToString(), "3");
