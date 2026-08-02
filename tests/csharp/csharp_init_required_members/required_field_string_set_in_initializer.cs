// vybe-test: csharp/csharp_init_required_members/required_field_string_set_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Header { public required string Name; }
var h = new Header { Name = "Content-Type" };
__Check((h.Name).ToString(), "Content-Type");
