// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_return_of_struct_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Widget{public int Id;} Widget w=new Widget(); w.Id=7; ref readonly int Get(ref Widget item)=>ref item.Id; __Check((Get(ref w)).ToString(), "7");
