// vybe-test: csharp/csharp_icomparable_sorting/custom_icomparer_reverses_natural_order
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Desc : System.Collections.Generic.IComparer<int> {
    public int Compare(int x, int y) => y.CompareTo(x);
}
var list = new System.Collections.Generic.List<int>{3,1,4,1,5};
list.Sort(new Desc());
__Check((list[0]).ToString(), "5");
