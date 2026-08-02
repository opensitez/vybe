// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_property_getter_still_reads
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class S{[Obsolete("old")] public int Value=>42;} __Check((new S().Value).ToString(), "42");
