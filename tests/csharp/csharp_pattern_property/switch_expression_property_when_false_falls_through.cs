// vybe-test: csharp/csharp_pattern_property/switch_expression_property_when_false_falls_through
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item { public int Q; } string Flag(object o)=>o switch{Item{Q:var q} when q>10=>"big",Item{Q:var q}=>"small",_=>"?"}; __Check((Flag(new Item{Q=3})).ToString(), "small");
