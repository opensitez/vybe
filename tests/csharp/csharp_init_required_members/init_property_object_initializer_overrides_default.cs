// vybe-test: csharp/csharp_init_required_members/init_property_object_initializer_overrides_default
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config { public int Port { get; init; } = 80; }
var c = new Config { Port = 443 };
__Check((c.Port).ToString(), "443");
