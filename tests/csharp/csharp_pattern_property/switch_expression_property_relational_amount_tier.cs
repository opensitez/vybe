// vybe-test: csharp/csharp_pattern_property/switch_expression_property_relational_amount_tier
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Bill { public int Amount; } string Tier(object o)=>o switch{Bill{Amount:>=100}=>"gold",Bill{Amount:>=50}=>"silver",_=>"bronze"}; __Check((Tier(new Bill{Amount=75})).ToString(), "silver");
