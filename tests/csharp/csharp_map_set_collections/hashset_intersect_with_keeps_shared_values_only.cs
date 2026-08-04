// vybe-test: csharp/csharp_map_set_collections/hashset_intersect_with_keeps_shared_values_only
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

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

using System.Collections.Generic; var left = new HashSet<int> { 1, 2, 3 }; left.IntersectWith(new[] { 2, 3, 4 }); foreach (var item in left) __P((item).ToString());
__Check("2\n3");
