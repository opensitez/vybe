// vybe-test: csharp/collections_advanced/hashset_intersect
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

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

var a = new HashSet<int> { 1, 2, 3, 4 };
var b = new HashSet<int> { 2, 4, 6 };
a.IntersectWith(b);
__P((a.Count).ToString());
__P((a.Contains(2)).ToString());
__P((a.Contains(4)).ToString());
__Check("2\nTrue\nTrue");
