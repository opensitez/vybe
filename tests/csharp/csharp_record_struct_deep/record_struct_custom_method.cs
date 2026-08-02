// vybe-test: csharp/csharp_record_struct_deep/record_struct_custom_method
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct V(int N){public int Twice()=>N*2;} __Check((new V(6).Twice()).ToString(), "12");
