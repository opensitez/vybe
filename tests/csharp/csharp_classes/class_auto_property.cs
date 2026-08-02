// vybe-test: csharp/csharp_classes/class_auto_property
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config {
    public string Name { get; set; }
    public int Value { get; set; }
}
var c = new Config();
c.Name = "test";
c.Value = 42;
__Check((c.Name).ToString(), "test");
__Check((c.Value).ToString(), "42");
