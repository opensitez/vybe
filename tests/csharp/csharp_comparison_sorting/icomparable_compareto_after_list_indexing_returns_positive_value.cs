// vybe-test: csharp/csharp_comparison_sorting/icomparable_compareto_after_list_indexing_returns_positive_value
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; class Rank : System.IComparable<Rank> { public int Value; public Rank(int value) { Value = value; } public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var list = new List<Rank> { new Rank(3), new Rank(1) }; __Check((list[0].CompareTo(list[1])).ToString(), "1");
