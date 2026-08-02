// vybe-test: csharp/csharp_pattern_property/is_property_pattern_nullable_int_null_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Holder { public int? Slot; } object o=new Holder{Slot=null}; __Check((o is Holder{Slot:null}).ToString(), "True");
