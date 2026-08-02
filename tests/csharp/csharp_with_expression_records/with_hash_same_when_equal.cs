// vybe-test: csharp/csharp_with_expression_records/with_hash_same_when_equal
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Key(int Id); var a=new Key(1); var b=a with{Id=1}; __Check((a.GetHashCode()==b.GetHashCode()).ToString(), "True");
