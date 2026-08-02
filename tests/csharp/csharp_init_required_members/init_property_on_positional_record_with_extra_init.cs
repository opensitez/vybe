// vybe-test: csharp/csharp_init_required_members/init_property_on_positional_record_with_extra_init
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User(string Name) { public int Age { get; init; } = 0; }
var u = new User("Bob") { Age = 30 };
__Check((u.Name).ToString(), "Bob"); __Check((u.Age).ToString(), "30");
