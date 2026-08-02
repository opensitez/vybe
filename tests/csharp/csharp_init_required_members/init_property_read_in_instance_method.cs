// vybe-test: csharp/csharp_init_required_members/init_property_read_in_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config { public int Port { get; init; } = 80; public int DoublePort() => Port * 2; }
var c = new Config { Port = 11 };
__Check((c.DoublePort()).ToString(), "22");
