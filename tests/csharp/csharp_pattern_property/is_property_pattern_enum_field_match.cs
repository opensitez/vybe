// vybe-test: csharp/csharp_pattern_property/is_property_pattern_enum_field_match
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color { Red, Green } class Paint { public Color Hue; } object o=new Paint{Hue=Color.Green}; __Check((o is Paint{Hue:Color.Green}).ToString(), "True");
