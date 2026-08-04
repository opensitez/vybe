// vybe-test: csharp/csharp_sorted_collections/sorted_set_intersect_with_keeps_sorted_overlap
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

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

using System.Collections.Generic; var a = new SortedSet<int> { 1, 2, 3, 4 }; a.IntersectWith(new[] { 3, 4, 5 }); __P((a.Count).ToString()); __P((a.Min).ToString());
__Check("2\n3");
