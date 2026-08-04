// vybe-test: csharp/csharp_linq_groupby_join/group_by_clusters_elements_by_key_and_counts_each_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

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

var words = new[] { "apple", "ant", "banana", "bear", "avocado" };
var groups = words
    .GroupBy(w => w[0])
    .OrderBy(g => g.Key)
    .Select(g => $"{g.Key}:{g.Count()}");
foreach (var s in groups) __P((s).ToString());
__Check("a:3\nb:2");
