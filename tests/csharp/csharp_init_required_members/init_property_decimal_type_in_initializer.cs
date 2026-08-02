// vybe-test: csharp/csharp_init_required_members/init_property_decimal_type_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Price { public decimal Amount { get; init; } }
var p = new Price { Amount = 19.99m };
__Check((p.Amount).ToString(), "19.99");
