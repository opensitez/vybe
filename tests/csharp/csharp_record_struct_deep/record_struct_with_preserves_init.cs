// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_preserves_init
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Pair{public int A{get;init;} public int B{get;init;}} var p=new Pair{A=1,B=2}; var q=p with{A=9}; __Check((q.B).ToString(), "2");
