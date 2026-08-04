// vybe-test: csharp/interfaces_generics/icomparable_implementation
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

class Temperature : IComparable<Temperature> {
    public double Degrees;
    public Temperature(double d) { Degrees = d; }
    public int CompareTo(Temperature other) {
        return Degrees.CompareTo(other.Degrees);
    }
    public override string ToString() { return Degrees + "°"; }
}
var temps = new List<Temperature> {
    new Temperature(100),
    new Temperature(37),
    new Temperature(0)
};
temps.Sort();
foreach (var t in temps) __P((t).ToString());
__Check("0°\n37°\n100°");
