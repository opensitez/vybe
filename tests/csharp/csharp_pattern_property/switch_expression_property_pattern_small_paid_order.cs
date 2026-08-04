// vybe-test: csharp/csharp_pattern_property/switch_expression_property_pattern_small_paid_order
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

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

class Order { public int Amount; public bool Paid; } string Label(object o)=>o switch{Order{Paid:true,Amount:>50}=>"big-paid",Order{Paid:true}=>"paid",_=>"open"}; __P((Label(new Order{Amount=10,Paid=true})).ToString());
__Check("paid");
