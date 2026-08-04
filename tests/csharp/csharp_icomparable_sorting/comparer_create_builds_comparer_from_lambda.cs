// vybe-test: csharp/csharp_icomparable_sorting/comparer_create_builds_comparer_from_lambda
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

var cmp = System.Collections.Generic.Comparer<string>.Create(
    (a,b) => a.Length.CompareTo(b.Length));
var list = new System.Collections.Generic.List<string>{"cc","aaa","b"};
list.Sort(cmp);
__P((list[0]).ToString());
__Check("b");
