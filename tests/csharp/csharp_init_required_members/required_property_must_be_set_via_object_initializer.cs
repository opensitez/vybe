// vybe-test: csharp/csharp_init_required_members/required_property_must_be_set_via_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Order { public required string Sku { get; set; } }
var o = new Order { Sku = "ABC" };
__Check((o.Sku).ToString(), "ABC");
