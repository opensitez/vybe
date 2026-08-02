// vybe-test: csharp/csharp_comparison_sorting/icomparable_direct_compareto_invocation_returns_positive_value
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Rank : System.IComparable<Rank> { public int Value; public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var left = new Rank(); left.Value = 3; var right = new Rank(); right.Value = 1; __Check((left.CompareTo(right)).ToString(), "1");
