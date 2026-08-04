// vybe-test: csharp/csharp_bcl_collections/read_only_list_view_reflects_backing_list_mutations
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

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

var backing = new System.Collections.Generic.List<int> { 1 };
var view = backing.AsReadOnly();
backing.Add(2);
__P((view.Count).ToString());
__Check("2");
