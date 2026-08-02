// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_var_syntax
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (x,y)=(5,7); __Check((x).ToString(), "5"); __Check((y).ToString(), "7");
