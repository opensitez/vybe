// vybe-test: csharp/csharp_init_required_members/init_property_bool_set_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Flags { public bool Enabled { get; init; } }
var f = new Flags { Enabled = true };
__Check((f.Enabled).ToString(), "True");
