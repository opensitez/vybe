// vybe-test: csharp/csharp_properties_accessors/validated_setter_rejects_negative_values
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var thermometer = new Thermometer();
thermometer.Celsius = -7;
__P((thermometer.Celsius).ToString());
thermometer.Celsius = 18;
__P((thermometer.Celsius).ToString());
__Check("0\n18");

class Thermometer {
    int celsius;
    public int Celsius {
        get { return celsius; }
        set { celsius = value < 0 ? 0 : value; }
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
