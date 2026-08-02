// vybe-test: csharp/csharp_properties/backing_field_used_in_custom_setter_logic
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Temperature {
    private double _celsius;
    public double Celsius {
        get => _celsius;
        set => _celsius = value < -273.15 ? -273.15 : value;
    }
}
var t = new Temperature(); t.Celsius = -300;
__Check((t.Celsius).ToString(), "-273.15");
