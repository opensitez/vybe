// vybe-test: csharp/csharp_with_expression_records/with_after_mutate_original
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Box{public int V{get;set;}} var a=new Box{V=1}; a.V=3; var b=a with{V=4}; __Check((b.V).ToString(), "4");
