// vybe-test: csharp/csharp_pattern_property/struct_property_pattern_on_value_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Vec2 { public int X; public int Y; } object o=new Vec2{X=2,Y=3}; __Check((o is Vec2{X:2,Y:3}).ToString(), "True");
