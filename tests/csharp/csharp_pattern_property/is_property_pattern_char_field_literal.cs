// vybe-test: csharp/csharp_pattern_property/is_property_pattern_char_field_literal
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Glyph { public char Ch; } object o=new Glyph{Ch='Z'}; __Check((o is Glyph{Ch:'Z'}).ToString(), "True");
