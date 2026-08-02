// vybe-test: csharp/csharp_pattern_property/is_property_pattern_rejects_wrong_int_field
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public int Value; } object o=new Box{Value=10}; __Check((o is Box{Value:11}).ToString(), "False");
