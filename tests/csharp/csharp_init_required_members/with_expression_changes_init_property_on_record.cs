// vybe-test: csharp/csharp_init_required_members/with_expression_changes_init_property_on_record
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config { public int Port { get; init; } = 80; }
var a = new Config();
var b = a with { Port = 9000 };
__Check((a.Port).ToString(), "80"); __Check((b.Port).ToString(), "9000");
