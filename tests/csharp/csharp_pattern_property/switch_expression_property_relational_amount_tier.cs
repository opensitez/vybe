// vybe-test: csharp/csharp_pattern_property/switch_expression_property_relational_amount_tier
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

class Bill { public int Amount; } string Tier(object o)=>o switch{Bill{Amount:>=100}=>"gold",Bill{Amount:>=50}=>"silver",_=>"bronze"}; __P((Tier(new Bill{Amount=75})).ToString());
__Check("silver");
