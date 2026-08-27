// vybe-test: csharp/csharp_oop/class_with_property_getset
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var t = new Temperature();
t.Celsius = 100;
__P((t.Celsius).ToString());
__P((t.Fahrenheit).ToString());
__Check("100\n212");

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
