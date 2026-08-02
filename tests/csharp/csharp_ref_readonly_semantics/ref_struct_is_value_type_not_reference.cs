// vybe-test: csharp/csharp_ref_readonly_semantics/ref_struct_is_value_type_not_reference
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

ref struct S{public int N;} var a=new S(); a.N=1; var b=a; b.N=2; __Check((a.N).ToString(), "1");
