// vybe-test: csharp/oop_advanced/sealed_class_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

sealed class Config {
    public string Name { get; set; }
    public Config(string n) { Name = n; }
}
var c = new Config("prod");
__Check((c.Name).ToString(), "prod");
