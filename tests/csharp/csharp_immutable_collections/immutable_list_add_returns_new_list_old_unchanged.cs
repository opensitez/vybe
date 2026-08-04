// vybe-test: csharp/csharp_immutable_collections/immutable_list_add_returns_new_list_old_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

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

var a=System.Collections.Immutable.ImmutableList<int>.Empty;
var b=a.Add(1).Add(2).Add(3);
__P((a.Count).ToString()); __P((b.Count).ToString());
__Check("0\n3");
