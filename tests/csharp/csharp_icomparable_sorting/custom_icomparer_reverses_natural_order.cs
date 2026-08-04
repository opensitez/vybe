// vybe-test: csharp/csharp_icomparable_sorting/custom_icomparer_reverses_natural_order
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

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

class Desc : System.Collections.Generic.IComparer<int> {
    public int Compare(int x, int y) => y.CompareTo(x);
}
var list = new System.Collections.Generic.List<int>{3,1,4,1,5};
list.Sort(new Desc());
__P((list[0]).ToString());
__Check("5");
