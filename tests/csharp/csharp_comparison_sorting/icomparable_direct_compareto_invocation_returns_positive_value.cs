// vybe-test: csharp/csharp_comparison_sorting/icomparable_direct_compareto_invocation_returns_positive_value
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

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

class Rank : System.IComparable<Rank> { public int Value; public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var left = new Rank(); left.Value = 3; var right = new Rank(); right.Value = 1; __P((left.CompareTo(right)).ToString());
__Check("1");
