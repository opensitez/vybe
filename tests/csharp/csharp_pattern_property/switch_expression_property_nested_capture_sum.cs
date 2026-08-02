// vybe-test: csharp/csharp_pattern_property/switch_expression_property_nested_capture_sum
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Inner { public int A; public int B; } class Wrap { public Inner Data; } int Sum(object o)=>o switch{Wrap{Data:{A:var a,B:var b}}=>a+b,_=>0}; __Check((Sum(new Wrap{Data=new Inner{A=6,B=7}})).ToString(), "13");
