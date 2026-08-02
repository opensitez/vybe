// vybe-test: csharp/csharp_init_required_members/required_enum_property_set_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Off, On }
class Switch { public required State Mode { get; set; } }
var s = new Switch { Mode = State.On };
__Check((s.Mode).ToString(), "On");
