// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_method_reads_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly ref struct Counter{public readonly int Value; public Counter(int v){Value=v;} public int Doubled()=>Value*2;} var c=new Counter(6); __Check((c.Doubled()).ToString(), "12");
