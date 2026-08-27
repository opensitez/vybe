// vybe-test: csharp/interfaces_generics/icomparable_implementation
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

var temps = new List<Temperature> {
    new Temperature(100),
    new Temperature(37),
    new Temperature(0)
}
;
temps.Sort();
foreach (var t in temps) __P((t).ToString());
__Check("0°\n37°\n100°");

class Temperature : IComparable<Temperature> {
    public double Degrees;
    public Temperature(double d) { Degrees = d; }
    public int CompareTo(Temperature other) {
        return Degrees.CompareTo(other.Degrees);
    }
    public override string ToString() { return Degrees + "°"; }
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
