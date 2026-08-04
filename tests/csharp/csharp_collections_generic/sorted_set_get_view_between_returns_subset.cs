// vybe-test: csharp/csharp_collections_generic/sorted_set_get_view_between_returns_subset
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

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

var s=new System.Collections.Generic.SortedSet<int>{1,2,3,4,5};
var view=s.GetViewBetween(2,4);
__P((view.Count).ToString());
__Check("3");
