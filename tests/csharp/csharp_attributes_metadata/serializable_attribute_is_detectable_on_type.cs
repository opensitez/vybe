// vybe-test: csharp/csharp_attributes_metadata/serializable_attribute_is_detectable_on_type
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Serializable] class Packet { } __Check((Attribute.IsDefined(typeof(Packet), typeof(SerializableAttribute))).ToString(), "True");
