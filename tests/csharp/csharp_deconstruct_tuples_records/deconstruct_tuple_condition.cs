// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_tuple_condition
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t=(2,5); var (x,y)=t; __Check((x<y).ToString(), "True");
