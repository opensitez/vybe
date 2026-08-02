// vybe-test: csharp/csharp_init_required_members/required_property_on_record_set_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Person { public required string Name { get; init; } }
var p = new Person { Name = "Rex" };
__Check((p.Name).ToString(), "Rex");
