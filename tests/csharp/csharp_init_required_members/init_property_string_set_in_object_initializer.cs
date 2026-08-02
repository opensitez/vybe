// vybe-test: csharp/csharp_init_required_members/init_property_string_set_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class User { public string Name { get; init; } = "guest"; }
var u = new User { Name = "Ada" };
__Check((u.Name).ToString(), "Ada");
