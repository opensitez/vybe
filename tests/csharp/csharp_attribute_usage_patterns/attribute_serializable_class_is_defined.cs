// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_class_is_defined
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Serializable] class Packet{} __Check((Attribute.IsDefined(typeof(Packet),typeof(SerializableAttribute))).ToString(), "True");
