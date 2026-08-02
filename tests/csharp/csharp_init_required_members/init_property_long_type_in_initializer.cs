// vybe-test: csharp/csharp_init_required_members/init_property_long_type_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Stats { public long Total { get; init; } }
var s = new Stats { Total = 10000000000L };
__Check((s.Total).ToString(), "10000000000");
