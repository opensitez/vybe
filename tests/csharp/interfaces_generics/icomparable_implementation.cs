// vybe-test: csharp/interfaces_generics/icomparable_implementation
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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
foreach (var t in temps) Console.WriteLine(t);
