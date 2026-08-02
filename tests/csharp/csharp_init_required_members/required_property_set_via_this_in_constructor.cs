// vybe-test: csharp/csharp_init_required_members/required_property_set_via_this_in_constructor
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Order { public required string Sku { get; set; } public Order(string sku) { Sku = sku; } }
var o = new Order("XYZ");
__Check((o.Sku).ToString(), "XYZ");
