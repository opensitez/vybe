// vybe-test: csharp/csharp_properties/auto_property_get_set_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Person { public string Name { get; set; } }
var p = new Person(); p.Name = "Alice";
__Check((p.Name).ToString(), "Alice");
