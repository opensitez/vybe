// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_overload_both_callable
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class S{[Obsolete("a")] public int Go()=>1; [Obsolete("b")] public int Go(int x)=>x;} __Check((new S().Go()).ToString(), "1"); __Check((new S().Go(5)).ToString(), "5");
