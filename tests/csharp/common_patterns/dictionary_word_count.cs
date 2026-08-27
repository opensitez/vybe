// vybe-test: csharp/common_patterns/dictionary_word_count
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

string text = "the cat sat on the mat the cat";
var words = text.Split(' ');
var counts = new Dictionary<string, int>();
foreach (var w in words) {
    if (counts.ContainsKey(w)) counts[w]++;
    else counts[w] = 1;
}
__P(("the: " + counts["the"]).ToString());
__P(("cat: " + counts["cat"]).ToString());
__P(("sat: " + counts["sat"]).ToString());
__Check("the: 3\ncat: 2\nsat: 1");

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
