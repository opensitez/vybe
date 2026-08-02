// vybe-test: csharp/csharp_init_required_members/init_property_enum_type_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Level { Low, High }
class Job { public Level Tier { get; init; } = Level.Low; }
var j = new Job { Tier = Level.High };
__Check((j.Tier).ToString(), "High");
