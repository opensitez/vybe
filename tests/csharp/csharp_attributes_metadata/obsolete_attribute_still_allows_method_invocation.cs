// vybe-test: csharp/csharp_attributes_metadata/obsolete_attribute_still_allows_method_invocation
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Service { [Obsolete("legacy")] public string Run() { return "ok"; } } __Check((new Service().Run()).ToString(), "ok");
