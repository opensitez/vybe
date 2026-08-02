// vybe-test: csharp/csharp_oop/class_auto_property_default
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config {
    public string Name { get; set; } = "default";
    public int Count { get; set; } = 0;
}
var c = new Config();
__Check((c.Name).ToString(), "default");
c.Name = "custom";
__Check((c.Name).ToString(), "custom");
