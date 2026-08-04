// vybe-test: csharp/csharp_comparison_sorting/icomparable_compareto_after_list_indexing_returns_positive_value
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

using System.Collections.Generic; class Rank : System.IComparable<Rank> { public int Value; public Rank(int value) { Value = value; } public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var list = new List<Rank> { new Rank(3), new Rank(1) }; __P((list[0].CompareTo(list[1])).ToString());
__Check("1");
