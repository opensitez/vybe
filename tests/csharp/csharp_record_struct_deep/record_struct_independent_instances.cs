// vybe-test: csharp/csharp_record_struct_deep/record_struct_independent_instances
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct V(int N); var a=new V(1); var b=new V(2); var c=a with{N=5}; __Check((b.N).ToString(), "2"); __Check((c.N).ToString(), "5");
