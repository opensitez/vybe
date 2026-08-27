// vybe-test: csharp/csharp_properties/backing_field_used_in_custom_setter_logic
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

using static __Harness;

var t = new Temperature();
t.Celsius = -300;
__P((t.Celsius).ToString());
__Check("-273.15");

class Temperature {
    private double _celsius;
    public double Celsius {
        get => _celsius;
        set => _celsius = value < -273.15 ? -273.15 : value;
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
