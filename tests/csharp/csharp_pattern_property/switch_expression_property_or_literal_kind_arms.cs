// vybe-test: csharp/csharp_pattern_property/switch_expression_property_or_literal_kind_arms
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Msg { public string Kind; } string Label(object o)=>o switch{Msg{Kind:"err" or "fail"}=>"bad",_=>"ok"}; __Check((Label(new Msg{Kind="fail"})).ToString(), "bad");
