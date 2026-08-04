// vybe-test: csharp/csharp_immutable_collections/immutable_dictionary_add_returns_new_dictionary
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

var d=System.Collections.Immutable.ImmutableDictionary<string,int>.Empty;
var d2=d.Add("a",1).Add("b",2);
__P((d.Count).ToString()); __P((d2["b"]).ToString());
__Check("0\n2");
