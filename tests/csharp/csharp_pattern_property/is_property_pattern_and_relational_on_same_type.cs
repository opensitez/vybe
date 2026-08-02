// vybe-test: csharp/csharp_pattern_property/is_property_pattern_and_relational_on_same_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Temp { public int C; } object o=new Temp{C=22}; __Check((o is Temp{C:>=20 and <=25}).ToString(), "True");
