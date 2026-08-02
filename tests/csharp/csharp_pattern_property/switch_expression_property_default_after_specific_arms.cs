// vybe-test: csharp/csharp_pattern_property/switch_expression_property_default_after_specific_arms
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token { public string Kind; } string Name(object o)=>o switch{Token{Kind:"add"}=>"plus",Token{Kind:"sub"}=>"minus",_=>"other"}; __Check((Name(new Token{Kind="mul"})).ToString(), "other");
