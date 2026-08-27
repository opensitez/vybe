// vybe-test: csharp/common_patterns/property_with_backing_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

var t = new Temperature();
t.Celsius = 100;
__P((t.Fahrenheit).ToString());
t.Celsius = -500;
__P((t.Celsius).ToString());
__Check("212\n-273.15");

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

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
