// vybe-test: csharp/csharp_pattern_property/switch_expression_property_pattern_big_paid_order
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Order { public int Amount; public bool Paid; } string Label(object o)=>o switch{Order{Paid:true,Amount:>50}=>"big-paid",Order{Paid:true}=>"paid",_=>"open"}; __Check((Label(new Order{Amount=100,Paid=true})).ToString(), "big-paid");
