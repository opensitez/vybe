// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_field_still_readable
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class S{[Obsolete("old")] public int N=6;} __Check((new S().N).ToString(), "6");
