// vybe-test: csharp/common_patterns/property_with_backing_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Temperature {
    private double celsius;
    public double Celsius {
        get { return celsius; }
        set {
            if (value < -273.15) celsius = -273.15;
            else celsius = value;
        }
    }
    public double Fahrenheit {
        get { return celsius * 9.0 / 5.0 + 32; }
    }
}
var t = new Temperature();
t.Celsius = 100;
__Check((t.Fahrenheit).ToString(), "212");
t.Celsius = -500;
__Check((t.Celsius).ToString(), "-273.15");
