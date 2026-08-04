// vybe-test: csharp/csharp_icomparable_sorting/custom_icomparable_used_by_array_sort
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

class Score : System.IComparable<Score> {
    public int Value;
    public int CompareTo(Score other) => Value.CompareTo(other.Value);
}
var scores = new[]{
    new Score{Value=5}, new Score{Value=1}, new Score{Value=3}
};
System.Array.Sort(scores);
__P((scores[0].Value).ToString());
__P((scores[2].Value).ToString());
__Check("1\n5");
