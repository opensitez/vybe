// vybe-test: csharp/csharp_init_required_members/init_property_default_preserved_across_two_instances
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config { public int Retries { get; init; } = 3; }
var a = new Config();
var b = new Config { Retries = 1 };
__Check((a.Retries).ToString(), "3"); __Check((b.Retries).ToString(), "1");
