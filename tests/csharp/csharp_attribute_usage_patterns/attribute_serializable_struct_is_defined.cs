// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_struct_is_defined
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Serializable] struct Point{} __Check((Attribute.IsDefined(typeof(Point),typeof(SerializableAttribute))).ToString(), "True");
