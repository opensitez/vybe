// vybe-test: csharp/csharp_icomparable_sorting/custom_icomparable_used_by_array_sort
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((scores[0].Value).ToString(), "1");
__Check((scores[2].Value).ToString(), "5");
