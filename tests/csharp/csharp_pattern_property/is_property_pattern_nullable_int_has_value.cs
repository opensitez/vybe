// vybe-test: csharp/csharp_pattern_property/is_property_pattern_nullable_int_has_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder { public int? Slot; } object o=new Holder{Slot=6}; __Check((o is Holder{Slot:6}).ToString(), "True");
