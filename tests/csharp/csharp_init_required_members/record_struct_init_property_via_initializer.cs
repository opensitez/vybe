// vybe-test: csharp/csharp_init_required_members/record_struct_init_property_via_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Tag { public string Name { get; init; } = "none"; }
var t = new Tag { Name = "alpha" };
__Check((t.Name).ToString(), "alpha");
