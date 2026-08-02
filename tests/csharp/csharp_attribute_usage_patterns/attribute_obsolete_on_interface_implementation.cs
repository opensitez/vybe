// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_on_interface_implementation
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; interface I{[Obsolete("old")] string Run();} class S:I{public string Run()=>"ok";} __Check((new S().Run()).ToString(), "ok");
