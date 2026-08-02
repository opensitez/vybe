// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_and_flags_on_different_types
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum F{A=1} [Obsolete("old")] class S{public int Use()=>(int)F.A;} __Check((new S().Use()).ToString(), "1");
