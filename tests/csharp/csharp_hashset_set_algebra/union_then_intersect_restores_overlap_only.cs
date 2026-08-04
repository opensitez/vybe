// vybe-test: csharp/csharp_hashset_set_algebra/union_then_intersect_restores_overlap_only
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

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

using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.UnionWith(new[] { 2, 3 }); a.IntersectWith(new[] { 2, 5 }); __P((a.Count).ToString()); __P((a.Contains(2)).ToString());
__Check("1\nTrue");
