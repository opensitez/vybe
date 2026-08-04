// vybe-test: csharp/csharp_bcl_collections/hashset_union_with_merges_distinct_elements_from_other_set
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

var left = new System.Collections.Generic.HashSet<int> { 1, 2 };
var right = new System.Collections.Generic.HashSet<int> { 2, 3 };
left.UnionWith(right);
__P((left.Count).ToString());
__Check("3");
