// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_inheritance_check
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Serializable] class Base{} class Derived:Base{} __Check((Attribute.IsDefined(typeof(Derived),typeof(SerializableAttribute))).ToString(), "False");
