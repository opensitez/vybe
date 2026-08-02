// vybe-test: csharp/csharp_init_required_members/init_property_default_used_when_initializer_omits_it
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config { public int Port { get; init; } = 8080; }
var c = new Config();
__Check((c.Port).ToString(), "8080");
