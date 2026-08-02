// vybe-test: csharp/csharp_properties/static_property_shares_value_across_instances
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Registry { public static int Count { get; set; } }
Registry.Count = 7;
__Check((Registry.Count).ToString(), "7");
