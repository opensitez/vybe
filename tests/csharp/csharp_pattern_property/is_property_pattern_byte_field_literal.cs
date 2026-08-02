// vybe-test: csharp/csharp_pattern_property/is_property_pattern_byte_field_literal
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Port { public byte P; } object o=new Port{P=80}; __Check((o is Port{P:80}).ToString(), "True");
