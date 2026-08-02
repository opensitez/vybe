// vybe-test: csharp/csharp_reflection_activation/property_info_sets_property_value_on_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public string Name { get; set; } } var box = new Box(); var prop = typeof(Box).GetProperty("Name"); prop.SetValue(box, "updated"); __Check((box.Name).ToString(), "updated");
