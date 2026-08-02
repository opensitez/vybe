// vybe-test: csharp/csharp_reflection_activation/property_info_reads_property_value_from_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public string Name { get; set; } = "pkg"; } var prop = typeof(Box).GetProperty("Name"); __Check((prop.GetValue(new Box())).ToString(), "pkg");
