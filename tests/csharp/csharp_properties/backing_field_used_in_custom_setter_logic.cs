// vybe-test: csharp/csharp_properties/backing_field_used_in_custom_setter_logic
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((t.Celsius).ToString());
__Check("-273.15");
