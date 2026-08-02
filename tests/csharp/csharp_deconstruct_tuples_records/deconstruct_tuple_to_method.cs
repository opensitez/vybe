// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_tuple_to_method
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Sum(int a,int b){__Check((a+b).ToString(), "5");} var (x,y)=(2,3); Sum(x,y);
