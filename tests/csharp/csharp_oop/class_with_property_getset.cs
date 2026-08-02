// vybe-test: csharp/csharp_oop/class_with_property_getset
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Temperature {
    private double _celsius;
    public double Celsius {
        get { return _celsius; }
        set { _celsius = value; }
    }
    public double Fahrenheit {
        get { return _celsius * 9 / 5 + 32; }
    }
}
var t = new Temperature();
t.Celsius = 100;
__Check((t.Celsius).ToString(), "100");
__Check((t.Fahrenheit).ToString(), "212");
