// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_mixed
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Mix(int N,string S); var (n,s)=new Mix(7,"x"); __Check((n).ToString(), "7"); __Check((s).ToString(), "x");
