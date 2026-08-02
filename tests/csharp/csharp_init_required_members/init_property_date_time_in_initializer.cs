// vybe-test: csharp/csharp_init_required_members/init_property_date_time_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Event { public System.DateTime When { get; init; } }
var e = new Event { When = new System.DateTime(2024, 1, 15) };
__Check((e.When.Year).ToString(), "2024");
