// vybe-test: csharp/csharp_pattern_property/switch_expression_property_when_guard_on_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair { public int A; public int B; } string Sign(object o)=>o switch{Pair{A:var x,B:var y} when x==y=>"eq",Pair{A:var x,B:var y}=>"neq",_=>"?"}; __Check((Sign(new Pair{A=3,B=3})).ToString(), "eq");
