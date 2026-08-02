// vybe-test: csharp/csharp_static_classes/static_readonly_field_set_in_static_constructor
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config {
    public static readonly string Version;
    static Config() { Version = "1.0"; }
}
__Check((Config.Version).ToString(), "1.0");
