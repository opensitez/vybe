// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_named_elements
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t=(X:2,Y:3); var (x,y)=t; __Check((x+y).ToString(), "5");
