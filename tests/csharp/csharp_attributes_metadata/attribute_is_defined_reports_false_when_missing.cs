// vybe-test: csharp/csharp_attributes_metadata/attribute_is_defined_reports_false_when_missing
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Plain { } __Check((Attribute.IsDefined(typeof(Plain), typeof(ObsoleteAttribute))).ToString(), "False");
