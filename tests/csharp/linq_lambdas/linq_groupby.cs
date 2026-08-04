// vybe-test: csharp/linq_lambdas/linq_groupby
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

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

var words = new List<string> { "apple", "ant", "banana", "avocado", "bat" };
var groups = words.GroupBy(w => w[0].ToString()).ToList();
foreach (var g in groups) {
    __P((g.Key + ": " + g.Count()).ToString());
}
__Check("a: 3\nb: 2");
