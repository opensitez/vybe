// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_after_with
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Pair(int A,int B); var q=(new Pair(1,2)) with{A=9}; var (a,b)=q; __Check((a).ToString(), "9"); __Check((b).ToString(), "2");
