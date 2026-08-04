// vybe-test: csharp/csharp_init_required_members/required_property_set_via_this_in_constructor
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

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

class Order { public required string Sku { get; set; } public Order(string sku) { Sku = sku; } }
var o = new Order("XYZ");
__P((o.Sku).ToString());
__Check("XYZ");
