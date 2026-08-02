// vybe-test: csharp/csharp_pattern_deconstruct/property_pattern_extracts_nested_property_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Order { public int Amount; public bool IsPaid; }
object o = new Order { Amount = 100, IsPaid = true };
var label = o switch {
    Order { IsPaid: true, Amount: > 50 } => "big paid",
    Order { IsPaid: true }               => "small paid",
    _                                    => "unpaid"
};
__Check((label).ToString(), "big paid");
