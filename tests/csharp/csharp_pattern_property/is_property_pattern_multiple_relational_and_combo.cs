// vybe-test: csharp/csharp_pattern_property/is_property_pattern_multiple_relational_and_combo
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Band { public int Lo; public int Hi; } object o=new Band{Lo=10,Hi=20}; __Check((o is Band{Lo:>=10 and <=10,Hi:>=20 and <=20}).ToString(), "True");
