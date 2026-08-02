// vybe-test: csharp/csharp_attributes_metadata/clscompliant_attribute_is_detectable_on_type
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [CLSCompliant(true)] class PublicApi { } __Check((Attribute.IsDefined(typeof(PublicApi), typeof(CLSCompliantAttribute))).ToString(), "True");
