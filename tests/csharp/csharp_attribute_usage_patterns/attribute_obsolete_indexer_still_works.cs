// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_indexer_still_works
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class S{[Obsolete("old")] public int this[int i]=>i*2;} __Check((new S()[4]).ToString(), "8");
