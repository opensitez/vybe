// vybe-test: csharp/csharp_reflection_activation/property_info_can_report_can_write_for_settable_property
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Box { public string Name { get; set; } } __Check((typeof(Box).GetProperty("Name").CanWrite).ToString(), "True");
