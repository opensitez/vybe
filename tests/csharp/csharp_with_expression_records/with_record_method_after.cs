// vybe-test: csharp/csharp_with_expression_records/with_record_method_after
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Counter(int N){public int Next()=>N+1;} var d=(new Counter(1)) with{N=5}; __Check((d.Next()).ToString(), "6");
