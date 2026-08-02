// vybe-test: csharp/csharp_properties_accessors/validated_setter_rejects_negative_values
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Thermometer {
    int celsius;
    public int Celsius {
        get { return celsius; }
        set { celsius = value < 0 ? 0 : value; }
    }
}
var thermometer = new Thermometer();
thermometer.Celsius = -7;
__Check((thermometer.Celsius).ToString(), "0");
thermometer.Celsius = 18;
__Check((thermometer.Celsius).ToString(), "18");
