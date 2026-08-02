// vybe-test: csharp/csharp_init_required_members/init_property_on_nominal_record_via_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Settings { public string Mode { get; init; } = "safe"; }
var s = new Settings { Mode = "fast" };
__Check((s.Mode).ToString(), "fast");
