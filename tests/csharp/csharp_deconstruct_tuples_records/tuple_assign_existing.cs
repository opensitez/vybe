// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_assign_existing
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=0,y=0; (x,y)=(9,1); __Check((x).ToString(), "9"); __Check((y).ToString(), "1");
