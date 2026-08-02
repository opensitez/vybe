// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_return_does_not_copy_large_struct_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Big{public int A; public int B; public int C;} Big item=new Big(); item.B=77; ref readonly int Read(ref Big target)=>ref target.B; __Check((Read(ref item)).ToString(), "77");
