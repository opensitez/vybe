// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_nested_type
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Outer{[Serializable] public class Inner{}} __Check((Attribute.IsDefined(typeof(Outer.Inner),typeof(SerializableAttribute))).ToString(), "True");
