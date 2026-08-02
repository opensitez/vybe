// vybe-test: csharp/csharp_init_required_members/init_property_and_public_field_in_same_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Form { public string Title { get; init; } public int Version; }
var f = new Form { Title = "main", Version = 2 };
__Check((f.Title).ToString(), "main"); __Check((f.Version).ToString(), "2");
