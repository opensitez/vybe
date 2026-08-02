// vybe-test: csharp/csharp_record_struct_deep/record_struct_static_factory
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct V(int N){public static V Zero()=>new V(0);} __Check((V.Zero().N).ToString(), "0");
