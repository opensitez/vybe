// vybe-test: csharp/csharp_properties/auto_property_with_default_value_initializer
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config { public int Timeout { get; set; } = 30; }
__Check((new Config().Timeout).ToString(), "30");
