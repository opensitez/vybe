// vybe-test: csharp/common_patterns/dictionary_grouping
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

var data = new List<string> { "apple", "banana", "avocado", "blueberry", "cherry" }
;
var grouped = new Dictionary<char, List<string>>();
foreach (var item in data) {
    char key = item[0];
    if (!grouped.ContainsKey(key)) grouped[key] = new List<string>();
    grouped[key].Add(item);
}
__P(("a: " + grouped['a'].Count).ToString());
__P(("b: " + grouped['b'].Count).ToString());
__P(("c: " + grouped['c'].Count).ToString());
__Check("a: 2\nb: 2\nc: 1");

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
