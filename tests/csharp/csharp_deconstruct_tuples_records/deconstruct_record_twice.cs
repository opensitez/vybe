// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_record_twice
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Pair(int A,int B); var p=new Pair(1,2); var (a,b)=p; var (c,d)=p; __Check((a+c).ToString(), "2"); __Check((b+d).ToString(), "4");
