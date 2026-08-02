// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_does_not_block_chained_calls
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class S{[Obsolete("old")] public string A()=>"a"; public string B()=>A()+"b";} __Check((new S().B()).ToString(), "ab");
