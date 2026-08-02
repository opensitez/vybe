// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_method_still_invokes
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class S{[Obsolete("old")] public string Run()=>"ok";} __Check((new S().Run()).ToString(), "ok");
