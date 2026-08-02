// vybe-test: csharp/csharp_init_required_members/init_property_mixed_with_settable_property
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item { public int Id { get; init; } public string Label { get; set; } = ""; }
var i = new Item { Id = 7 };
i.Label = "tool";
__Check((i.Id).ToString(), "7"); __Check((i.Label).ToString(), "tool");
