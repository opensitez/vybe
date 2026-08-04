// vybe-test: csharp/csharp_pattern_deconstruct/property_pattern_extracts_nested_property_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((label).ToString());
__Check("big paid");
