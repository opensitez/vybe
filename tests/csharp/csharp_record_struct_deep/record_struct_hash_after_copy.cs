// vybe-test: csharp/csharp_record_struct_deep/record_struct_hash_after_copy
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Key(int Id); var a=new Key(9); var b=a; __Check((a.GetHashCode()==b.GetHashCode()).ToString(), "True");
